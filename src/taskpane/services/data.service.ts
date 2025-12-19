import { Injectable } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import initSqlJs, { Database, SqlJsStatic } from "sql.js";

/* global Excel */

const CUSTOM_XML_ELEMENT = "proofPanelData";

@Injectable({ providedIn: "root" })
export class DataService {
  private sqlPromise?: Promise<SqlJsStatic>;
  private db?: Database;

  constructor(private logger: NGXLogger) {}

  async hasDatabase(): Promise<boolean> {
    if (!this.isExcelAvailable()) {
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
    this.ensureExcelAvailable();

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

    this.ensureExcelAvailable();
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
    this.ensureExcelAvailable();

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
    if (!this.isExcelAvailable()) {
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

  private isExcelAvailable(): boolean {
    this.logger.debug("isExcelAvailable: checking Excel global");
    return typeof Excel !== "undefined";
  }

  private extractBase64(xml: string): string | null {
    this.logger.debug("extractBase64: extracting base64 payload from XML");
    const match = new RegExp(`<${CUSTOM_XML_ELEMENT}>([^<]*)</${CUSTOM_XML_ELEMENT}>`).exec(xml);
    return match?.[1]?.trim() || null;
  }

  private uint8ArrayToBase64(bytes: Uint8Array): string {
    this.logger.debug("uint8ArrayToBase64: converting bytes to base64");
    let binary = "";
    bytes.forEach((b) => {
      binary += String.fromCharCode(b);
    });
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
}
