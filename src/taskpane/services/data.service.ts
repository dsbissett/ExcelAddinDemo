import { Injectable } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import { Database } from "sql.js";
import { sha256Hex } from "./utils/binary-encoding";
import { createDocumentId, createPdfPartUri, createRelationshipId } from "./utils/identifiers";
import { ExcelHostService } from "./excel-host.service";
import { WorkbookPackageService } from "./workbook-package.service";
import { DataFileRepository, DataFileRecord, SavedFilePart, SaveFilePartOptions } from "./data-file-repository.service";

/* global Excel */

export { DataFileRecord, SavedFilePart, SaveFilePartOptions } from "./data-file-repository.service";

const PDF_CONTENT_TYPE = "application/pdf";
const ENABLE_PACKAGE_PART_WRITES = false;

@Injectable({ providedIn: "root" })
export class DataService {
  private db?: Database;
  private readonly requiredTables = ["Pages", "Cells", "PolygonData"];

  constructor(
    private logger: NGXLogger,
    private excelHost: ExcelHostService,
    private workbookPackage: WorkbookPackageService,
    private dataFileRepo: DataFileRepository,
  ) {}

  // ── Database Lifecycle ──────────────────────────────────────────────

  async hasDatabase(): Promise<boolean> {
    if (!(await this.excelHost.waitForOfficeReady())) {
      this.logger.warn("hasDatabase: Excel not available");
      return false;
    }

    let exists = false;
    await Excel.run(async (context) => {
      const part = await this.excelHost.findDataPart(context);
      exists = Boolean(part);
    });

    this.logger.info(`hasDatabase: ${exists ? "found existing database" : "no database found"}`);
    return exists;
  }

  async getDatabaseState(
    requiredTables: string[] = this.requiredTables,
  ): Promise<{ hasDatabase: boolean; hasData: boolean; missingRequiredTables: string[] }> {
    if (!(await this.excelHost.waitForOfficeReady())) {
      this.logger.warn("getDatabaseState: Excel not available");
      return { hasDatabase: false, hasData: false, missingRequiredTables: [...requiredTables] };
    }

    const existing = await this.tryLoadFromWorkbook();
    const hasDatabase = Boolean(existing);
    const missingRequiredTables = existing ? this.getMissingTables(existing, requiredTables) : [...requiredTables];
    const hasData = existing ? this.hasUserData(existing) : false;

    if (existing && !this.db) {
      this.db = existing;
    }

    this.logger.info(
      `getDatabaseState: hasDatabase=${hasDatabase} hasData=${hasData} missingTables=${missingRequiredTables.join(",")}`,
    );
    return { hasDatabase, hasData, missingRequiredTables };
  }

  async seedDatabase(sqlText: string): Promise<void> {
    this.logger.info("seedDatabase: seeding database with provided SQL");
    const database = await this.loadOrCreate();
    database.run(sqlText);
    await this.saveDatabase();
    this.logger.info("seedDatabase: seeding completed and database saved");
  }

  async execute(sqlText: string): Promise<Array<{ columns: string[]; values: unknown[][] }>> {
    const start = Date.now();
    this.logger.info("execute: starting SQL query");
    const isReadOnly = this.isReadOnlyQuery(sqlText);
    const database = await this.loadOrCreate();

    try {
      const results = database.exec(sqlText).map((result) => ({
        columns: result.columns,
        values: result.values,
      }));
      this.logQueryCompletion(start, isReadOnly);
      return results;
    } catch (error) {
      this.logger.error(`execute: query failed after ${Date.now() - start}ms`, error);
      throw error;
    }
  }

  async loadOrCreate(): Promise<Database> {
    this.logger.info("loadOrCreate: ensure Excel available and load existing db if present");
    await this.excelHost.ensureExcelReady();

    if (!this.db) {
      const existing = await this.tryLoadFromWorkbook();
      if (existing) {
        this.logger.info("loadOrCreate: loaded database from customXml part");
        this.db = existing;
      } else {
        this.logger.info("loadOrCreate: no database found, creating new database");
        this.db = await this.createEmptyDatabase();
        await this.saveDatabase();
      }
    }

    return this.db;
  }

  async updateDatabase(mutator: (database: Database) => void | Promise<void>): Promise<Database> {
    this.logger.info("updateDatabase: loading database and applying mutator");
    const database = await this.loadOrCreate();
    await mutator(database);
    await this.saveDatabase();
    this.logger.info("updateDatabase: mutation applied and database saved");
    return database;
  }

  async saveDatabase(): Promise<void> {
    if (!this.db) {
      throw new Error("Database has not been initialized.");
    }
    await this.excelHost.ensureExcelReady();
    await this.excelHost.saveCustomXmlPayload(this.db);
  }

  async deleteDatabase(): Promise<void> {
    this.logger.info("deleteDatabase: clearing in-memory db and removing customXml part");
    this.db = undefined;
    await this.excelHost.ensureExcelReady();
    await this.excelHost.deleteCustomXmlPayload();
  }

  // ── Data File Operations ────────────────────────────────────────────

  async saveFilePart(fileName: string, payload: Uint8Array, options?: SaveFilePartOptions): Promise<SavedFilePart> {
    this.logger.info(`saveFilePart: saving file ${fileName} to workbook package part`);
    await this.excelHost.ensureExcelReady();

    const pdfBytes = payload instanceof Uint8Array ? payload : new Uint8Array(payload);
    const partUri = createPdfPartUri(fileName);
    const relationshipId = createRelationshipId();
    const contentHash = await sha256Hex(pdfBytes);
    const createdUtc = new Date().toISOString();

    if (ENABLE_PACKAGE_PART_WRITES) {
      await this.writeFilePartToPackage(pdfBytes, partUri, relationshipId, options);
    } else {
      this.logger.debug("saveFilePart: package part writes disabled; payload will be persisted in sqlite blob storage");
    }

    this.logger.info(`saveFilePart: prepared ${fileName} with part uri ${partUri}`);
    return {
      documentId: createDocumentId(),
      fileName,
      contentType: PDF_CONTENT_TYPE,
      partUri,
      relationshipId,
      contentHash,
      version: 1,
      createdUtc,
      updatedUtc: createdUtc,
    };
  }

  async beginWorkbookPackageBatch(): Promise<void> {
    if (!ENABLE_PACKAGE_PART_WRITES) return;
    await this.excelHost.ensureExcelReady();
    await this.workbookPackage.beginBatch();
  }

  async commitWorkbookPackageBatch(): Promise<void> {
    if (!ENABLE_PACKAGE_PART_WRITES) return;
    await this.workbookPackage.commitBatch();
  }

  rollbackWorkbookPackageBatch(): void {
    if (!ENABLE_PACKAGE_PART_WRITES) return;
    this.workbookPackage.rollbackBatch();
  }

  async recordDataFile(record: DataFileRecord): Promise<void> {
    this.logger.info(`recordDataFile: recording file ${record.fileName} in database`);
    await this.recordDataFiles([record]);
  }

  async recordDataFiles(records: DataFileRecord[]): Promise<void> {
    if (!records.length) return;

    this.logger.info(`recordDataFiles: recording ${records.length} file(s) in database`);
    const database = await this.loadOrCreate();
    this.dataFileRepo.ensureSchema(database);
    this.dataFileRepo.upsertAll(database, records);
    await this.saveDatabase();
  }

  async deleteDataFile(partUri: string): Promise<void> {
    this.logger.info(`deleteDataFile: removing package part ${partUri} and record`);
    await this.excelHost.ensureExcelReady();
    const database = await this.loadOrCreate();
    this.dataFileRepo.ensureSchema(database);

    if (ENABLE_PACKAGE_PART_WRITES) {
      await this.deleteFilePartFromPackage(database, partUri);
    }

    this.dataFileRepo.deleteByPartUri(database, partUri);
    await this.saveDatabase();
    this.logger.info(`deleteDataFile: removed package part and record for ${partUri}`);
  }

  async getDataFiles(): Promise<DataFileRecord[]> {
    this.logger.info("getDataFiles: loading stored data files");
    const database = await this.loadOrCreate();
    this.dataFileRepo.ensureSchema(database);
    return this.dataFileRepo.queryAll(database);
  }

  async loadFilePart(partUri: string): Promise<Uint8Array | null> {
    this.logger.info(`loadFilePart: loading package part ${partUri}`);
    await this.excelHost.ensureExcelReady();

    const database = await this.loadOrCreate();
    this.dataFileRepo.ensureSchema(database);
    const blobPayload = this.dataFileRepo.loadPayloadByPartUri(database, partUri);
    if (blobPayload) return blobPayload;

    if (ENABLE_PACKAGE_PART_WRITES) {
      return this.loadFilePartFromPackage(partUri);
    }

    this.logger.warn(`loadFilePart: no payload found for ${partUri}`);
    return null;
  }

  // ── Table Inspection ────────────────────────────────────────────────

  async listTables(): Promise<string[]> {
    this.logger.info("listTables: loading table list");
    const database = await this.loadOrCreate();
    return this.getUserTables(database);
  }

  async getTableRowCount(tableName: string): Promise<number> {
    this.logger.info(`getTableRowCount: counting rows for ${tableName}`);
    const database = await this.loadOrCreate();
    const safeTable = this.escapeIdentifier(tableName);
    try {
      const result = database.exec(`SELECT COUNT(*) AS rowCount FROM ${safeTable};`);
      const value = result?.[0]?.values?.[0]?.[0];
      return typeof value === "number" ? value : Number(value) || 0;
    } catch (error) {
      this.logger.error(`getTableRowCount: failed for ${tableName}`, error);
      return 0;
    }
  }

  async previewTable(tableName: string, limit: number = 100): Promise<{ columns: string[]; values: unknown[][] }> {
    this.logger.info(`previewTable: loading preview for ${tableName}`);
    const database = await this.loadOrCreate();
    const safeTable = this.escapeIdentifier(tableName);
    const safeLimit = Math.max(1, Math.min(Math.floor(limit), 500));
    try {
      const result = database.exec(`SELECT * FROM ${safeTable} LIMIT ${safeLimit};`);
      const first = result?.[0];
      return { columns: first?.columns ?? [], values: first?.values ?? [] };
    } catch (error) {
      this.logger.error(`previewTable: failed for ${tableName}`, error);
      return { columns: [], values: [] };
    }
  }

  // ── Private Helpers ─────────────────────────────────────────────────

  private async tryLoadFromWorkbook(): Promise<Database | null> {
    this.logger.debug("tryLoadFromWorkbook: checking for existing customXml database");
    if (!(await this.excelHost.waitForOfficeReady())) {
      this.logger.warn("tryLoadFromWorkbook: Excel not available");
      return null;
    }

    const sql = await this.excelHost.loadSqlEngine();
    const xmlPayload = await this.excelHost.readCustomXmlPayload();
    if (!xmlPayload) return null;

    const bytes = this.excelHost.extractDatabaseBytes(xmlPayload);
    if (!bytes) return null;

    try {
      return new sql.Database(bytes);
    } catch (error) {
      this.logger.error("tryLoadFromWorkbook: failed to load database from customXml part", error);
      return null;
    }
  }

  private async createEmptyDatabase(): Promise<Database> {
    this.logger.debug("createEmptyDatabase: creating new empty database");
    const sql = await this.excelHost.loadSqlEngine();
    return new sql.Database();
  }

  private isReadOnlyQuery(sqlText: string): boolean {
    const withoutComments = sqlText.replace(/\/\*[\s\S]*?\*\//g, " ").replace(/--.*$/gm, " ");
    const normalized = withoutComments.trim().replace(/^[();\s]+/, "");
    return Boolean(normalized.match(/^(with|select|pragma|explain)\b/i));
  }

  private logQueryCompletion(startTime: number, isReadOnly: boolean): void {
    const duration = Date.now() - startTime;
    if (isReadOnly) {
      this.logger.info(`execute: query completed in ${duration}ms (read-only; no save)`);
      return;
    }

    this.logger.info(`execute: query completed in ${duration}ms (writes detected; save queued)`);
    void this.saveDatabase()
      .then(() => this.logger.info("execute: database saved after write"))
      .catch((saveError) => this.logger.error("execute: failed to save database after write", saveError));
  }

  private hasUserData(database: Database): boolean {
    const tables = this.getUserTables(database);
    return tables.some((table) => this.tableHasRows(database, table));
  }

  private tableHasRows(database: Database, table: string): boolean {
    try {
      const result = database.exec(`SELECT EXISTS(SELECT 1 FROM "${table}" LIMIT 1) AS hasData;`);
      const value = result?.[0]?.values?.[0]?.[0];
      return value === 1 || value === true;
    } catch (error) {
      this.logger.warn(`tableHasRows: failed to check table ${table}`, error);
      return false;
    }
  }

  private getMissingTables(database: Database, requiredTables: string[]): string[] {
    const existingTables = this.getUserTables(database).map((name) => name.toLowerCase());
    return requiredTables.filter((table) => !existingTables.includes(table.toLowerCase()));
  }

  private getUserTables(database: Database): string[] {
    try {
      const result = database.exec(
        "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%';",
      );
      return result?.[0]?.values?.map((row) => String(row[0])) ?? [];
    } catch (error) {
      this.logger.warn("getUserTables: failed to list user tables", error);
      return [];
    }
  }

  private escapeIdentifier(identifier: string): string {
    const trimmed = identifier?.trim() ?? "";
    const escaped = trimmed.replace(/"/g, '""');
    return `"${escaped}"`;
  }

  private async writeFilePartToPackage(
    pdfBytes: Uint8Array,
    partUri: string,
    relationshipId: string,
    options?: SaveFilePartOptions,
  ): Promise<void> {
    const shouldDefer = options?.deferWorkbookWrite ?? this.workbookPackage.hasPendingBatch();
    const entries = await this.workbookPackage.getEntriesForWrite(shouldDefer);
    entries[this.workbookPackage.partUriToZipPath(partUri)] = pdfBytes;
    this.workbookPackage.addOrUpdateContentTypeOverride(entries, partUri, PDF_CONTENT_TYPE);
    this.workbookPackage.addOrUpdateRelationship(entries, partUri, relationshipId);

    if (shouldDefer) {
      this.workbookPackage.markPendingDirty();
    } else {
      await this.workbookPackage.writePackage(entries);
    }
  }

  private async deleteFilePartFromPackage(database: Database, partUri: string): Promise<void> {
    const relationshipId = this.dataFileRepo.findRelationshipId(database, partUri);
    const entries = await this.workbookPackage.readPackage();
    await this.workbookPackage.removePackagePart(entries, partUri, relationshipId);
  }

  private async loadFilePartFromPackage(partUri: string): Promise<Uint8Array | null> {
    const entries = await this.workbookPackage.readPackage();
    const entry = entries[this.workbookPackage.partUriToZipPath(partUri)];
    return entry ? new Uint8Array(entry) : null;
  }
}
