import { Injectable } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import initSqlJs, { Database, SqlJsStatic } from "sql.js";

/* global Excel */

const CUSTOM_XML_ELEMENT = "proofPanelData";
const DATA_FILE_ELEMENT_PREFIX = "dataFile";

export interface DataFileRecord {
  fileName: string;
  xmlPartName: string;
  rawFileSize: number;
  thumbnailPng: Uint8Array | null;
  thumbnailWidth: number | null;
  thumbnailHeight: number | null;
  thumbnailMimeType: string;
  createdUtc: string;
}

@Injectable({ providedIn: "root" })
export class DataService {
  private sqlPromise?: Promise<SqlJsStatic>;
  private db?: Database;
  private readonly requiredTables = ["Pages", "Cells", "PolygonData"];
  private officeReadyPromise?: Promise<unknown>;

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
    payload: Uint8Array | string,
  ): Promise<{ xmlPartName: string; createdUtc: string }> {
    this.logger.info(`saveFilePart: saving file ${fileName} to customXml`);
    await this.ensureExcelReady();

    const base64Payload =
      payload instanceof Uint8Array ? this.uint8ArrayToBase64(payload) : String(payload).trim();
    const xmlPartName = this.createDataFileElementName(fileName);
    const createdUtc = new Date().toISOString();
    const xml = `<?xml version="1.0" encoding="UTF-8"?><${xmlPartName} name="${this.escapeXmlAttribute(
      fileName,
    )}" created="${this.escapeXmlAttribute(createdUtc)}">${base64Payload}</${xmlPartName}>`;

    await Excel.run(async (context) => {
      context.workbook.customXmlParts.add(xml);
      await context.sync();
    });

    this.logger.info(`saveFilePart: saved ${fileName} to customXml with element ${xmlPartName}`);
    return { xmlPartName, createdUtc };
  }

  async recordDataFile(record: DataFileRecord): Promise<void> {
    this.logger.info(`recordDataFile: recording file ${record.fileName} in database`);
    const database = await this.loadOrCreate();
    this.ensureDataFilesTable(database);
    database.run(
      `INSERT OR REPLACE INTO DataFiles
        (FileName, XmlPartName, RawFileSize, ThumbnailPng, ThumbnailWidth, ThumbnailHeight, ThumbnailMimeType, CreatedUtc)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?);`,
      [
        record.fileName,
        record.xmlPartName,
        record.rawFileSize,
        record.thumbnailPng,
        record.thumbnailWidth,
        record.thumbnailHeight,
        record.thumbnailMimeType,
        record.createdUtc,
      ],
    );
    await this.saveDatabase();
  }

  async deleteDataFile(xmlPartName: string): Promise<void> {
    this.logger.info(`deleteDataFile: removing xml part ${xmlPartName} and record`);
    await this.ensureExcelReady();

    // Delete the customXml part
    await Excel.run(async (context) => {
      const parts = context.workbook.customXmlParts;
      parts.load("items");
      await context.sync();

      const xmlResults = parts.items.map((part) => part.getXml());
      await context.sync();

      for (let i = 0; i < parts.items.length; i += 1) {
        const xml = xmlResults[i].value ?? "";
        if (xml.includes(`<${xmlPartName}`)) {
          parts.items[i].delete();
          break;
        }
      }
      await context.sync();
    });

    // Delete the database record
    const database = await this.loadOrCreate();
    this.ensureDataFilesTable(database);
    database.run(`DELETE FROM DataFiles WHERE XmlPartName = ?;`, [xmlPartName]);
    await this.saveDatabase();
    this.logger.info(`deleteDataFile: removed xml part and record for ${xmlPartName}`);
  }

  async getDataFiles(): Promise<DataFileRecord[]> {
    this.logger.info("getDataFiles: loading stored data files");
    const database = await this.loadOrCreate();
    this.ensureDataFilesTable(database);

    try {
      const result = database.exec(
        `SELECT FileName, XmlPartName, RawFileSize, ThumbnailPng, ThumbnailWidth, ThumbnailHeight, ThumbnailMimeType, CreatedUtc
         FROM DataFiles
         ORDER BY datetime(CreatedUtc) DESC;`,
      );
      const rows = result?.[0]?.values ?? [];
      return rows.map((row) => {
        const mime = row[6];
        const thumbnailMimeType =
          typeof mime === "string" && mime.trim() ? mime : "image/png";
        const thumbnailPng = row[3] instanceof Uint8Array ? (row[3] as Uint8Array) : null;
        return {
          fileName: String(row[0]),
          xmlPartName: String(row[1]),
          rawFileSize: Number(row[2]),
          thumbnailPng,
          thumbnailWidth: row[4] !== null ? Number(row[4]) : null,
          thumbnailHeight: row[5] !== null ? Number(row[5]) : null,
          thumbnailMimeType,
          createdUtc: String(row[7]),
        };
      });
    } catch (error) {
      this.logger.error("getDataFiles: failed to query DataFiles", error);
      return [];
    }
  }

  async loadFilePart(xmlPartName: string): Promise<string | null> {
    this.logger.info(`loadFilePart: loading customXml part ${xmlPartName}`);
    await this.ensureExcelReady();

    let payload: string | null = null;
    let matchedXml: string | null = null;

    await Excel.run(async (context) => {
      const parts = context.workbook.customXmlParts;
      parts.load("items");
      await context.sync();

      const xmlResults = parts.items.map((part) => part.getXml());
      await context.sync();

      for (let i = 0; i < parts.items.length; i += 1) {
        const xml = xmlResults[i].value ?? "";
        if (xml.includes(`<${xmlPartName}`)) {
          matchedXml = xml;
          payload = this.extractBase64(xml, xmlPartName);
          if (payload) {
            break;
          }
        }
      }
    });

    if (!payload && matchedXml) {
      const decodedXml = this.tryDecodeBase64Xml(matchedXml);
      if (decodedXml) {
        payload = this.extractBase64(decodedXml, xmlPartName);
        if (!payload) {
          this.logger.warn(`loadFilePart: decoded base64 xml but still no payload for ${xmlPartName}`);
        }
      }
    }

    if (!payload) {
      this.logger.warn(`loadFilePart: no payload found for ${xmlPartName}`);
    }

    return payload;
  }

  private ensureDataFilesTable(database: Database): void {
    database.run(
      `CREATE TABLE IF NOT EXISTS DataFiles (
        DataFileID          INTEGER PRIMARY KEY AUTOINCREMENT,
        FileName            TEXT NOT NULL,
        XmlPartName         TEXT NOT NULL,
        RawFileSize         INTEGER NOT NULL,
        ThumbnailPng        BLOB NULL,
        ThumbnailWidth      INTEGER NULL,
        ThumbnailHeight     INTEGER NULL,
        ThumbnailMimeType   TEXT NOT NULL DEFAULT 'image/png',
        CreatedUtc          TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
        UNIQUE (FileName),
        CHECK (length(trim(FileName)) > 0),
        CHECK (length(trim(XmlPartName)) > 0),
        CHECK (RawFileSize >= 0),
        CHECK (ThumbnailWidth IS NULL OR ThumbnailWidth > 0),
        CHECK (ThumbnailHeight IS NULL OR ThumbnailHeight > 0)
      );`,
    );
    database.run(`CREATE INDEX IF NOT EXISTS IX_DataFiles_XmlPartName ON DataFiles(XmlPartName);`);
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

  private ensureExcelAvailable(): void {
    this.logger.debug("ensureExcelAvailable: verifying Excel runtime is present");
    if (!this.isExcelAvailable()) {
      throw new Error("Excel is not available. Connect to Excel before accessing the database.");
    }
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

  private tryDecodeBase64Xml(input: string): string | null {
    const trimmed = input.trim();
    // If it already looks like XML, don't attempt base64 decode.
    if (!trimmed || trimmed.includes("<")) {
      return null;
    }

    try {
      const decoded = atob(trimmed);
      if (decoded.includes("<")) {
        this.logger.debug("tryDecodeBase64Xml: decoded xml from base64 wrapper");
        return decoded;
      }
    } catch (error) {
      this.logger.debug("tryDecodeBase64Xml: base64 decode failed", error);
    }

    return null;
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

  private createDataFileElementName(fileName: string): string {
    const safeName = fileName
      .replace(/[^a-zA-Z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase() || "file";
    const suffix = Math.random().toString(36).slice(2, 8);
    return `${DATA_FILE_ELEMENT_PREFIX}-${safeName}-${suffix}`;
  }

  private escapeXmlAttribute(value: string): string {
    return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
}
