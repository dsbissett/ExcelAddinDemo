import { Injectable } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import initSqlJs, { Database, SqlJsStatic } from "sql.js";

/* global Excel */

const CUSTOM_XML_ELEMENT = "proofPanelData";
const PDF_PARTS_FOLDER = "pdfs";
const PDF_CONTENT_TYPE = "application/pdf";
const PACKAGE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const ENABLE_PACKAGE_PART_WRITES = false;

export interface DataFileRecord {
  documentId: string;
  fileName: string;
  contentType: string;
  partUri: string;
  relationshipId: string | null;
  contentHash: string;
  version: number;
  pdfPayload?: Uint8Array | null;
  rawFileSize: number;
  thumbnailPng: Uint8Array | null;
  thumbnailWidth: number | null;
  thumbnailHeight: number | null;
  thumbnailMimeType: string;
  createdUtc: string;
  updatedUtc: string;
}

export interface SavedFilePart {
  documentId: string;
  fileName: string;
  contentType: string;
  partUri: string;
  relationshipId: string | null;
  contentHash: string;
  version: number;
  createdUtc: string;
  updatedUtc: string;
}

export interface SaveFilePartOptions {
  deferWorkbookWrite?: boolean;
}

@Injectable({ providedIn: "root" })
export class DataService {
  private sqlPromise?: Promise<SqlJsStatic>;
  private db?: Database;
  private readonly requiredTables = ["Pages", "Cells", "PolygonData"];
  private officeReadyPromise?: Promise<unknown>;
  private pendingWorkbookPackageEntries?: Record<string, Uint8Array>;
  private isPendingWorkbookPackageDirty = false;

  constructor(private logger: NGXLogger) {}

  async hasDatabase(): Promise<boolean> {
    if (!(await this.waitForOfficeReady())) {
      this.logger.warn("hasDatabase: Excel not available");
      return false;
    }

    let exists = false;
    await Excel.run(async (context) => {
      const part = await this.findDataPart(context);
      exists = Boolean(part);
    });

    this.logger.info(`hasDatabase: ${exists ? "found existing database" : "no database found"}`);
    return exists;
  }

  async getDatabaseState(
    requiredTables: string[] = this.requiredTables,
  ): Promise<{ hasDatabase: boolean; hasData: boolean; missingRequiredTables: string[] }> {
    if (!(await this.waitForOfficeReady())) {
      this.logger.warn("getDatabaseState: Excel not available");
      return { hasDatabase: false, hasData: false, missingRequiredTables: [...requiredTables] };
    }

    const existing = await this.tryLoadFromWorkbook();
    const hasDatabase = Boolean(existing);
    const missingRequiredTables = existing
      ? this.getMissingTables(existing, requiredTables)
      : [...requiredTables];
    const hasData = existing ? this.hasUserData(existing) : false;

    if (existing && !this.db) {
      // Cache the loaded database for subsequent operations.
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
      const duration = Date.now() - start;

      if (!isReadOnly) {
        this.logger.info(`execute: query completed in ${duration}ms (writes detected; save queued)`);
        void this.saveDatabase()
          .then(() => this.logger.info("execute: database saved after write"))
          .catch((saveError) => this.logger.error("execute: failed to save database after write", saveError));
      } else {
        this.logger.info(`execute: query completed in ${duration}ms (read-only; no save)`);
      }

      return results;
    } catch (error) {
      const duration = Date.now() - start;
      this.logger.error(`execute: query failed after ${duration}ms`, error);
      throw error;
    }
  }

  private isReadOnlyQuery(sqlText: string): boolean {
    const withoutComments = sqlText
      .replace(/\/\*[\s\S]*?\*\//g, " ")
      .replace(/--.*$/gm, " ");
    const normalized = withoutComments.trim().replace(/^[();\s]+/, "");
    const match = normalized.match(/^(with|select|pragma|explain)\b/i);
    return Boolean(match);
  }

  async loadOrCreate(): Promise<Database> {
    this.logger.info("loadOrCreate: ensure Excel available and load existing db if present");
    await this.ensureExcelReady();

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
    this.logger.info("saveDatabase: exporting database and writing to customXml");
    if (!this.db) {
      throw new Error("Database has not been initialized.");
    }

    await this.ensureExcelReady();
    const payload = this.uint8ArrayToBase64(this.db.export());
    const xml = `<?xml version="1.0" encoding="UTF-8"?><${CUSTOM_XML_ELEMENT}>${payload}</${CUSTOM_XML_ELEMENT}>`;

    await Excel.run(async (context) => {
      const existingPart = await this.findDataPart(context);
      if (existingPart) {
        existingPart.delete();
      }

      context.workbook.customXmlParts.add(xml);
      await context.sync();
    });
    this.logger.info("saveDatabase: database saved to customXml");
  }

  async deleteDatabase(): Promise<void> {
    this.logger.info("deleteDatabase: clearing in-memory db and removing customXml part");
    this.db = undefined;
    await this.ensureExcelReady();

    await Excel.run(async (context) => {
      const existingPart = await this.findDataPart(context);
      if (existingPart) {
        existingPart.delete();
        await context.sync();
      }
    });
  }

  private async tryLoadFromWorkbook(): Promise<Database | null> {
    this.logger.debug("tryLoadFromWorkbook: checking for existing customXml database");
    if (!(await this.waitForOfficeReady())) {
      this.logger.warn("tryLoadFromWorkbook: Excel not available");
      return null;
    }

    const sql = await this.loadSql();
    let xmlPayload: string | null = null;

    await Excel.run(async (context) => {
      const dataPart = await this.findDataPart(context);
      if (!dataPart) {
        this.logger.debug("tryLoadFromWorkbook: no data part found");
        return;
      }

      const xmlResult = dataPart.getXml();
      await context.sync();
      xmlPayload = xmlResult.value ?? null;
    });

    if (!xmlPayload) {
      this.logger.debug("tryLoadFromWorkbook: no XML payload to load");
      return null;
    }

    const base64 = this.extractBase64(xmlPayload);
    if (!base64) {
      this.logger.debug("tryLoadFromWorkbook: base64 payload missing in XML");
      return null;
    }

    try {
      const bytes = this.base64ToUint8Array(base64);
      return new sql.Database(bytes);
    } catch (error) {
      this.logger.error("tryLoadFromWorkbook: failed to load database from customXml part", error);
      return null;
    }
  }

  private async createEmptyDatabase(): Promise<Database> {
    this.logger.debug("createEmptyDatabase: creating new empty database");
    const sql = await this.loadSql();
    return new sql.Database();
  }

  private hasUserData(database: Database): boolean {
    const tables = this.getUserTables(database);
    for (const table of tables) {
      try {
        const result = database.exec(`SELECT EXISTS(SELECT 1 FROM "${table}" LIMIT 1) AS hasData;`);
        const value = result?.[0]?.values?.[0]?.[0];
        if (value === 1 || value === true) {
          return true;
        }
      } catch (error) {
        this.logger.warn(`hasUserData: failed to check table ${table}`, error);
      }
    }

    return false;
  }

  private getMissingTables(database: Database, requiredTables: string[]): string[] {
    const existingTables = this.getUserTables(database).map((name) => name.toLowerCase());
    return requiredTables.filter((table) => !existingTables.includes(table.toLowerCase()));
  }

  async saveFilePart(
    fileName: string,
    payload: Uint8Array,
    options?: SaveFilePartOptions,
  ): Promise<SavedFilePart> {
    this.logger.info(`saveFilePart: saving file ${fileName} to workbook package part`);
    await this.ensureExcelReady();
    const pdfBytes = payload instanceof Uint8Array ? payload : new Uint8Array(payload);
    const documentId = this.createDocumentId();
    const partUri = this.createPdfPartUri(fileName);
    const relationshipId = this.createRelationshipId();
    const createdUtc = new Date().toISOString();
    const updatedUtc = createdUtc;
    const contentHash = await this.sha256Hex(pdfBytes);

    if (ENABLE_PACKAGE_PART_WRITES) {
      const shouldDeferWrite = options?.deferWorkbookWrite ?? this.hasPendingWorkbookPackageBatch();
      const packageEntries = await this.getWorkbookPackageForWrite(shouldDeferWrite);
      const zipPath = this.partUriToZipPath(partUri);
      packageEntries[zipPath] = pdfBytes;
      this.addOrUpdateContentTypeOverride(packageEntries, partUri, PDF_CONTENT_TYPE);
      this.addOrUpdateWorkbookRelationship(packageEntries, partUri, relationshipId);
      if (shouldDeferWrite) {
        this.isPendingWorkbookPackageDirty = true;
      } else {
        await this.writeWorkbookPackage(packageEntries);
      }
    } else {
      this.logger.debug("saveFilePart: package part writes disabled; payload will be persisted in sqlite blob storage");
    }
    this.logger.info(
      `saveFilePart: prepared ${fileName} with part uri ${partUri} (package writes ${
        ENABLE_PACKAGE_PART_WRITES ? "enabled" : "disabled"
      })`,
    );
    return {
      documentId,
      fileName,
      contentType: PDF_CONTENT_TYPE,
      partUri,
      relationshipId,
      contentHash,
      version: 1,
      createdUtc,
      updatedUtc,
    };
  }

  async beginWorkbookPackageBatch(): Promise<void> {
    if (!ENABLE_PACKAGE_PART_WRITES) {
      return;
    }
    await this.ensureExcelReady();
    if (!this.pendingWorkbookPackageEntries) {
      this.pendingWorkbookPackageEntries = await this.readWorkbookPackage();
      this.isPendingWorkbookPackageDirty = false;
      this.logger.info("beginWorkbookPackageBatch: initialized staged workbook package");
    }
  }

  async commitWorkbookPackageBatch(): Promise<void> {
    if (!ENABLE_PACKAGE_PART_WRITES) {
      return;
    }
    if (!this.pendingWorkbookPackageEntries) {
      return;
    }

    if (this.isPendingWorkbookPackageDirty) {
      await this.writeWorkbookPackage(this.pendingWorkbookPackageEntries);
      this.logger.info("commitWorkbookPackageBatch: committed staged workbook package");
    } else {
      this.logger.info("commitWorkbookPackageBatch: no staged workbook changes to commit");
    }

    this.pendingWorkbookPackageEntries = undefined;
    this.isPendingWorkbookPackageDirty = false;
  }

  rollbackWorkbookPackageBatch(): void {
    if (!ENABLE_PACKAGE_PART_WRITES) {
      return;
    }
    this.pendingWorkbookPackageEntries = undefined;
    this.isPendingWorkbookPackageDirty = false;
    this.logger.info("rollbackWorkbookPackageBatch: cleared staged workbook package");
  }

  async recordDataFile(record: DataFileRecord): Promise<void> {
    this.logger.info(`recordDataFile: recording file ${record.fileName} in database`);
    await this.recordDataFiles([record]);
  }

  async recordDataFiles(records: DataFileRecord[]): Promise<void> {
    if (!records.length) {
      return;
    }

    this.logger.info(`recordDataFiles: recording ${records.length} file(s) in database`);
    const database = await this.loadOrCreate();
    this.ensureDataFilesTable(database);

    try {
      database.run("BEGIN TRANSACTION;");
      for (const record of records) {
        this.upsertDataFileRecord(database, record);
      }
      database.run("COMMIT;");
    } catch (error) {
      try {
        database.run("ROLLBACK;");
      } catch (rollbackError) {
        this.logger.warn("recordDataFiles: rollback failed", rollbackError);
      }
      throw error;
    }

    await this.saveDatabase();
  }

  async deleteDataFile(partUri: string): Promise<void> {
    this.logger.info(`deleteDataFile: removing package part ${partUri} and record`);
    await this.ensureExcelReady();
    const database = await this.loadOrCreate();
    this.ensureDataFilesTable(database);
    if (ENABLE_PACKAGE_PART_WRITES) {
      const lookup = database.exec(`SELECT RelationshipId FROM DataFiles WHERE PartUri = ? LIMIT 1;`, [partUri]);
      const relationshipId =
        lookup?.[0]?.values?.[0]?.[0] !== null && lookup?.[0]?.values?.[0]?.[0] !== undefined
          ? String(lookup[0].values[0][0])
          : null;

      const packageEntries = await this.readWorkbookPackage();
      delete packageEntries[this.partUriToZipPath(partUri)];
      this.removeContentTypeOverride(packageEntries, partUri);
      this.removeWorkbookRelationship(packageEntries, partUri, relationshipId);
      await this.writeWorkbookPackage(packageEntries);
    }

    database.run(`DELETE FROM DataFiles WHERE PartUri = ?;`, [partUri]);
    await this.saveDatabase();
    this.logger.info(`deleteDataFile: removed package part and record for ${partUri}`);
  }

  async getDataFiles(): Promise<DataFileRecord[]> {
    this.logger.info("getDataFiles: loading stored data files");
    const database = await this.loadOrCreate();
    this.ensureDataFilesTable(database);

    try {
      const result = database.exec(
        `SELECT DocumentID, FileName, ContentType, PartUri, RelationshipId, ContentHash, Version, RawFileSize,
                ThumbnailPng, ThumbnailWidth, ThumbnailHeight, ThumbnailMimeType, CreatedUtc, UpdatedUtc
         FROM DataFiles
         WHERE PartUri IS NOT NULL AND length(trim(PartUri)) > 0
         ORDER BY datetime(CreatedUtc) DESC;`,
      );
      const rows = result?.[0]?.values ?? [];
      return rows.map((row) => {
        const mime = row[11];
        const thumbnailMimeType =
          typeof mime === "string" && mime.trim() ? mime : "image/png";
        const thumbnailPng = row[8] instanceof Uint8Array ? (row[8] as Uint8Array) : null;
        return {
          documentId: String(row[0]),
          fileName: String(row[1]),
          contentType: String(row[2] ?? PDF_CONTENT_TYPE),
          partUri: String(row[3]),
          relationshipId: row[4] !== null ? String(row[4]) : null,
          contentHash: String(row[5] ?? ""),
          version: Number(row[6] ?? 1),
          rawFileSize: Number(row[7]),
          thumbnailPng,
          thumbnailWidth: row[9] !== null ? Number(row[9]) : null,
          thumbnailHeight: row[10] !== null ? Number(row[10]) : null,
          thumbnailMimeType,
          createdUtc: String(row[12]),
          updatedUtc: String(row[13] ?? row[12]),
        };
      });
    } catch (error) {
      this.logger.error("getDataFiles: failed to query DataFiles", error);
      return [];
    }
  }

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

  async previewTable(
    tableName: string,
    limit: number = 100,
  ): Promise<{ columns: string[]; values: unknown[][] }> {
    this.logger.info(`previewTable: loading preview for ${tableName}`);
    const database = await this.loadOrCreate();
    const safeTable = this.escapeIdentifier(tableName);
    const safeLimit = Math.max(1, Math.min(Math.floor(limit), 500));
    try {
      const result = database.exec(`SELECT * FROM ${safeTable} LIMIT ${safeLimit};`);
      const first = result?.[0];
      return {
        columns: first?.columns ?? [],
        values: first?.values ?? [],
      };
    } catch (error) {
      this.logger.error(`previewTable: failed for ${tableName}`, error);
      return { columns: [], values: [] };
    }
  }

  async loadFilePart(partUri: string): Promise<Uint8Array | null> {
    this.logger.info(`loadFilePart: loading package part ${partUri}`);
    await this.ensureExcelReady();

    const database = await this.loadOrCreate();
    this.ensureDataFilesTable(database);
    try {
      const row = database.exec(`SELECT PdfPayload FROM DataFiles WHERE PartUri = ? LIMIT 1;`, [partUri]);
      const payload = row?.[0]?.values?.[0]?.[0];
      if (payload instanceof Uint8Array && payload.length > 0) {
        return new Uint8Array(payload);
      }
    } catch (error) {
      this.logger.warn("loadFilePart: failed to read sqlite PdfPayload blob", error);
    }

    if (ENABLE_PACKAGE_PART_WRITES) {
      const packageEntries = await this.readWorkbookPackage();
      const entry = packageEntries[this.partUriToZipPath(partUri)];
      if (entry) {
        return new Uint8Array(entry);
      }
    }

    this.logger.warn(`loadFilePart: no payload found for ${partUri}`);
    return null;
  }

  private upsertDataFileRecord(database: Database, record: DataFileRecord): void {
    database.run(
      `INSERT OR REPLACE INTO DataFiles
        (DocumentID, FileName, ContentType, PartUri, RelationshipId, ContentHash, Version, PdfPayload, RawFileSize, ThumbnailPng, ThumbnailWidth, ThumbnailHeight, ThumbnailMimeType, CreatedUtc, UpdatedUtc)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);`,
      [
        record.documentId,
        record.fileName,
        record.contentType,
        record.partUri,
        record.relationshipId,
        record.contentHash,
        record.version,
        record.pdfPayload ?? null,
        record.rawFileSize,
        record.thumbnailPng,
        record.thumbnailWidth,
        record.thumbnailHeight,
        record.thumbnailMimeType,
        record.createdUtc,
        record.updatedUtc,
      ],
    );
  }

  private ensureDataFilesTable(database: Database): void {
    database.run(
      `CREATE TABLE IF NOT EXISTS DataFiles (
        DataFileID          INTEGER PRIMARY KEY AUTOINCREMENT,
        DocumentID          TEXT NOT NULL,
        FileName            TEXT NOT NULL,
        ContentType         TEXT NOT NULL DEFAULT 'application/pdf',
        PartUri             TEXT NOT NULL,
        RelationshipId      TEXT NULL,
        ContentHash         TEXT NOT NULL DEFAULT '',
        Version             INTEGER NOT NULL DEFAULT 1,
        PdfPayload          BLOB NULL,
        RawFileSize         INTEGER NOT NULL,
        ThumbnailPng        BLOB NULL,
        ThumbnailWidth      INTEGER NULL,
        ThumbnailHeight     INTEGER NULL,
        ThumbnailMimeType   TEXT NOT NULL DEFAULT 'image/png',
        CreatedUtc          TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
        UpdatedUtc          TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
        UNIQUE (DocumentID),
        UNIQUE (PartUri),
        CHECK (length(trim(FileName)) > 0),
        CHECK (length(trim(DocumentID)) > 0),
        CHECK (length(trim(ContentType)) > 0),
        CHECK (length(trim(PartUri)) > 0),
        CHECK (RawFileSize >= 0),
        CHECK (Version >= 1),
        CHECK (ThumbnailWidth IS NULL OR ThumbnailWidth > 0),
        CHECK (ThumbnailHeight IS NULL OR ThumbnailHeight > 0)
      );`,
    );
    this.ensureDataFilesColumns(database);
    database.run(`CREATE INDEX IF NOT EXISTS IX_DataFiles_PartUri ON DataFiles(PartUri);`);
    database.run(`CREATE INDEX IF NOT EXISTS IX_DataFiles_RelationshipId ON DataFiles(RelationshipId);`);
    database.run(`CREATE INDEX IF NOT EXISTS IX_DataFiles_DocumentId ON DataFiles(DocumentID);`);
  }

  private ensureDataFilesColumns(database: Database): void {
    const info = database.exec("PRAGMA table_info(DataFiles);");
    const existing = new Set(
      (info?.[0]?.values ?? []).map((row) => String(row[1]).toLowerCase()),
    );
    const ensureColumn = (name: string, ddl: string): void => {
      if (!existing.has(name.toLowerCase())) {
        database.run(`ALTER TABLE DataFiles ADD COLUMN ${ddl};`);
      }
    };

    ensureColumn("DocumentID", "DocumentID TEXT");
    ensureColumn("ContentType", "ContentType TEXT DEFAULT 'application/pdf'");
    ensureColumn("PartUri", "PartUri TEXT");
    ensureColumn("RelationshipId", "RelationshipId TEXT");
    ensureColumn("ContentHash", "ContentHash TEXT DEFAULT ''");
    ensureColumn("Version", "Version INTEGER DEFAULT 1");
    ensureColumn("PdfPayload", "PdfPayload BLOB");
    ensureColumn("UpdatedUtc", "UpdatedUtc TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))");

    database.run(
      `UPDATE DataFiles
       SET DocumentID = lower(hex(randomblob(16)))
       WHERE DocumentID IS NULL OR length(trim(DocumentID)) = 0;`,
    );
    database.run(
      `UPDATE DataFiles
       SET ContentType = 'application/pdf'
       WHERE ContentType IS NULL OR length(trim(ContentType)) = 0;`,
    );
    database.run(
      `UPDATE DataFiles
       SET ContentHash = ''
       WHERE ContentHash IS NULL;`,
    );
    database.run(
      `UPDATE DataFiles
       SET Version = 1
       WHERE Version IS NULL OR Version < 1;`,
    );
    database.run(
      `UPDATE DataFiles
       SET UpdatedUtc = COALESCE(NULLIF(trim(UpdatedUtc), ''), CreatedUtc, strftime('%Y-%m-%dT%H:%M:%fZ','now'))
       WHERE UpdatedUtc IS NULL OR length(trim(UpdatedUtc)) = 0;`,
    );
  }

  private createDocumentId(): string {
    const cryptoApi = globalThis.crypto;
    if (cryptoApi?.randomUUID) {
      return cryptoApi.randomUUID();
    }
    return `doc-${Date.now().toString(36)}-${this.randomToken(12)}`;
  }

  private createPdfPartUri(fileName: string): string {
    const stripped = fileName.replace(/\.pdf$/i, "");
    const safeName = stripped
      .replace(/[^a-zA-Z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase() || "document";
    return `/${PDF_PARTS_FOLDER}/${safeName}-${this.randomToken(8)}.pdf`;
  }

  private createRelationshipId(): string {
    return `rIdPdf${this.randomToken(10)}`;
  }

  private randomToken(length: number): string {
    const value = Math.random().toString(36).slice(2);
    if (value.length >= length) {
      return value.slice(0, length);
    }
    return `${value}${Math.random().toString(36).slice(2)}`.slice(0, length);
  }

  private normalizePartUri(partUri: string): string {
    const trimmed = String(partUri ?? "").trim().replace(/\\/g, "/");
    if (!trimmed) {
      throw new Error("Part URI cannot be empty.");
    }
    return trimmed.startsWith("/") ? trimmed : `/${trimmed}`;
  }

  private partUriToZipPath(partUri: string): string {
    return this.normalizePartUri(partUri).replace(/^\//, "");
  }

  private partUriToWorkbookRelationshipTarget(partUri: string): string {
    return `../${this.partUriToZipPath(partUri)}`;
  }

  private addOrUpdateContentTypeOverride(
    packageEntries: Record<string, Uint8Array>,
    partUri: string,
    contentType: string,
  ): void {
    const contentTypesPath = "[Content_Types].xml";
    const xml = packageEntries[contentTypesPath];
    if (!xml) {
      throw new Error("Workbook package is missing [Content_Types].xml.");
    }

    const doc = this.parseXmlDocument(strFromU8(xml));
    const normalizedPartUri = this.normalizePartUri(partUri);
    const overrides = Array.from(doc.getElementsByTagName("Override"));
    let override = overrides.find((item) => item.getAttribute("PartName") === normalizedPartUri);
    if (!override) {
      override = doc.createElementNS(doc.documentElement.namespaceURI, "Override");
      doc.documentElement.appendChild(override);
    }
    override.setAttribute("PartName", normalizedPartUri);
    override.setAttribute("ContentType", contentType);
    packageEntries[contentTypesPath] = strToU8(this.serializeXmlDocument(doc));
  }

  private removeContentTypeOverride(packageEntries: Record<string, Uint8Array>, partUri: string): void {
    const contentTypesPath = "[Content_Types].xml";
    const xml = packageEntries[contentTypesPath];
    if (!xml) {
      return;
    }
    const doc = this.parseXmlDocument(strFromU8(xml));
    const normalizedPartUri = this.normalizePartUri(partUri);
    const overrides = Array.from(doc.getElementsByTagName("Override"));
    for (const override of overrides) {
      if (override.getAttribute("PartName") === normalizedPartUri) {
        override.parentNode?.removeChild(override);
      }
    }
    packageEntries[contentTypesPath] = strToU8(this.serializeXmlDocument(doc));
  }

  private addOrUpdateWorkbookRelationship(
    packageEntries: Record<string, Uint8Array>,
    partUri: string,
    relationshipId: string,
  ): void {
    const relsPath = "xl/_rels/workbook.xml.rels";
    const xml = packageEntries[relsPath]
      ? strFromU8(packageEntries[relsPath])
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const doc = this.parseXmlDocument(xml);
    const relationships = Array.from(doc.getElementsByTagName("Relationship"));
    const target = this.partUriToWorkbookRelationshipTarget(partUri);
    let relationship = relationships.find((item) => item.getAttribute("Id") === relationshipId);
    if (!relationship) {
      relationship = doc.createElementNS(doc.documentElement.namespaceURI, "Relationship");
      doc.documentElement.appendChild(relationship);
    }

    relationship.setAttribute("Id", relationshipId);
    relationship.setAttribute("Type", PACKAGE_RELATIONSHIP_TYPE);
    relationship.setAttribute("Target", target);
    packageEntries[relsPath] = strToU8(this.serializeXmlDocument(doc));
  }

  private removeWorkbookRelationship(
    packageEntries: Record<string, Uint8Array>,
    partUri: string,
    relationshipId?: string | null,
  ): void {
    const relsPath = "xl/_rels/workbook.xml.rels";
    const xml = packageEntries[relsPath];
    if (!xml) {
      return;
    }
    const doc = this.parseXmlDocument(strFromU8(xml));
    const target = this.partUriToWorkbookRelationshipTarget(partUri);
    const relationships = Array.from(doc.getElementsByTagName("Relationship"));
    for (const relationship of relationships) {
      const byId = Boolean(relationshipId) && relationship.getAttribute("Id") === relationshipId;
      const byTarget = relationship.getAttribute("Target") === target;
      if (byId || byTarget) {
        relationship.parentNode?.removeChild(relationship);
      }
    }
    packageEntries[relsPath] = strToU8(this.serializeXmlDocument(doc));
  }

  private parseXmlDocument(xmlText: string): Document {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, "application/xml");
    if (doc.querySelector("parsererror")) {
      throw new Error("Unable to parse workbook package XML.");
    }
    return doc;
  }

  private serializeXmlDocument(doc: Document): string {
    const serializer = new XMLSerializer();
    return serializer.serializeToString(doc);
  }

  private async sha256Hex(bytes: Uint8Array): Promise<string> {
    const cryptoApi = globalThis.crypto;
    if (!cryptoApi?.subtle) {
      throw new Error("Web Crypto API is unavailable; cannot compute PDF hash.");
    }
    const digest = await cryptoApi.subtle.digest("SHA-256", bytes);
    return Array.from(new Uint8Array(digest))
      .map((value) => value.toString(16).padStart(2, "0"))
      .join("");
  }

  private async readWorkbookPackage(): Promise<Record<string, Uint8Array>> {
    const compressed = await this.getCompressedWorkbookBytes();
    return unzipSync(compressed);
  }

  private hasPendingWorkbookPackageBatch(): boolean {
    return Boolean(this.pendingWorkbookPackageEntries);
  }

  private async getWorkbookPackageForWrite(deferWorkbookWrite: boolean): Promise<Record<string, Uint8Array>> {
    if (deferWorkbookWrite) {
      if (!this.pendingWorkbookPackageEntries) {
        this.pendingWorkbookPackageEntries = await this.readWorkbookPackage();
        this.isPendingWorkbookPackageDirty = false;
      }
      return this.pendingWorkbookPackageEntries;
    }
    return this.readWorkbookPackage();
  }

  private async writeWorkbookPackage(packageEntries: Record<string, Uint8Array>): Promise<void> {
    const compressed = zipSync(packageEntries, { level: 6 });
    await Excel.createWorkbook(this.uint8ArrayToBase64(compressed));
  }

  private async getCompressedWorkbookBytes(): Promise<Uint8Array> {
    const file = await this.getCompressedFileHandle();
    try {
      const slices: Uint8Array[] = [];
      for (let index = 0; index < file.sliceCount; index += 1) {
        const slice = await this.getOfficeFileSlice(file, index);
        slices.push(this.normalizeSliceData(slice.data));
      }
      const totalLength = slices.reduce((sum, part) => sum + part.length, 0);
      const output = new Uint8Array(totalLength);
      let offset = 0;
      for (const part of slices) {
        output.set(part, offset);
        offset += part.length;
      }
      return output;
    } finally {
      await this.closeOfficeFileHandle(file);
    }
  }

  private getCompressedFileHandle(): Promise<Office.File> {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 1024 * 1024 },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
            return;
          }
          reject(
            new Error(
              result.error?.message ?? "Unable to read workbook package from the Office host.",
            ),
          );
        },
      );
    });
  }

  private getOfficeFileSlice(file: Office.File, index: number): Promise<Office.Slice> {
    return new Promise((resolve, reject) => {
      file.getSliceAsync(index, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
          return;
        }
        reject(
          new Error(
            result.error?.message ?? `Unable to read workbook package slice at index ${index}.`,
          ),
        );
      });
    });
  }

  private closeOfficeFileHandle(file: Office.File): Promise<void> {
    return new Promise((resolve) => {
      file.closeAsync(() => resolve());
    });
  }

  private normalizeSliceData(data: unknown): Uint8Array {
    if (data instanceof Uint8Array) {
      return data;
    }
    if (data instanceof ArrayBuffer) {
      return new Uint8Array(data);
    }
    if (ArrayBuffer.isView(data)) {
      return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    }
    if (Array.isArray(data)) {
      return Uint8Array.from(data as number[]);
    }
    throw new Error("Unexpected slice format when reading workbook package.");
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

  private async findDataPart(context: Excel.RequestContext): Promise<Excel.CustomXmlPart | null> {
    this.logger.debug("findDataPart: searching customXml parts for proofPanelData");
    const parts = context.workbook.customXmlParts;
    parts.load("items");
    await context.sync();

    const xmlResults = parts.items.map((part) => part.getXml());
    await context.sync();

    for (let i = 0; i < parts.items.length; i += 1) {
      const xml = xmlResults[i].value ?? "";
      if (xml.includes(`<${CUSTOM_XML_ELEMENT}`)) {
        this.logger.debug("findDataPart: found matching customXml part");
        return parts.items[i];
      }
    }

    this.logger.debug("findDataPart: no matching customXml part found");
    return null;
  }

  private async loadSql(): Promise<SqlJsStatic> {
    this.logger.debug("loadSql: loading sql.js");
    if (!this.sqlPromise) {
      this.logger.debug("loadSql: initializing sql.js");
      this.sqlPromise = initSqlJs({
        locateFile: (file) => file,
      });
    }

    return this.sqlPromise;
  }

  private async waitForOfficeReady(): Promise<boolean> {
    if (!this.officeReadyPromise) {
      if (typeof Office === "undefined" || !Office.onReady) {
        return false;
      }
      this.officeReadyPromise = Office.onReady();
    }

    try {
      await this.officeReadyPromise;
      return true;
    } catch (error) {
      this.logger.error("waitForOfficeReady: Office.onReady failed", error);
      this.officeReadyPromise = undefined;
      return false;
    }
  }

  private async ensureExcelReady(): Promise<void> {
    const ready = await this.waitForOfficeReady();
    if (!ready || !this.isExcelAvailable()) {
      throw new Error("Excel is not available. Connect to Excel before accessing the database.");
    }
  }

  private isExcelAvailable(): boolean {
    this.logger.debug("isExcelAvailable: checking Excel global");
    return typeof Excel !== "undefined";
  }

  private extractBase64(xml: string, elementName: string = CUSTOM_XML_ELEMENT): string | null {
    this.logger.debug("extractBase64: extracting base64 payload from XML");
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(xml, "application/xml");
      const parserError = doc.querySelector("parsererror");
      if (parserError) {
        this.logger.warn("extractBase64: XML parse error, falling back to string search");
      } else {
        const byName = doc.querySelector(elementName);
        const payload = (byName ?? doc.documentElement)?.textContent?.trim() ?? "";
        if (payload) {
          return payload;
        }
      }
    } catch (error) {
      this.logger.warn("extractBase64: DOMParser failed, falling back to string search", error);
    }

    const openTag = `<${elementName}`;
    const start = xml.indexOf(openTag);
    if (start === -1) {
      return null;
    }
    const startContent = xml.indexOf(">", start);
    const end = xml.indexOf(`</${elementName}>`, startContent);
    if (startContent === -1 || end === -1 || end <= startContent) {
      return null;
    }
    const payload = xml.slice(startContent + 1, end).trim();
    return payload || null;
  }

  private uint8ArrayToBase64(bytes: Uint8Array): string {
    this.logger.debug("uint8ArrayToBase64: converting bytes to base64");
    const chunkSize = 0x8000;
    let binary = "";
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode(...chunk);
    }
    return btoa(binary);
  }

  private base64ToUint8Array(base64: string): Uint8Array {
    this.logger.debug("base64ToUint8Array: converting base64 to bytes");
    const binary = atob(base64);
    const buffer = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      buffer[i] = binary.charCodeAt(i);
    }
    return buffer;
  }

  private escapeIdentifier(identifier: string): string {
    const trimmed = identifier?.trim() ?? "";
    const escaped = trimmed.replace(/"/g, "\"\"");
    return `"${escaped}"`;
  }
}
