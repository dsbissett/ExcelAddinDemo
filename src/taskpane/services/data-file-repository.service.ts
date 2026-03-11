import { Injectable } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import { Database } from "sql.js";

const PDF_CONTENT_TYPE = "application/pdf";

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

interface ColumnMigration {
  name: string;
  ddl: string;
}

interface ColumnBackfill {
  sql: string;
}

const COLUMN_MIGRATIONS: ColumnMigration[] = [
  { name: "DocumentID", ddl: "DocumentID TEXT" },
  { name: "ContentType", ddl: "ContentType TEXT DEFAULT 'application/pdf'" },
  { name: "PartUri", ddl: "PartUri TEXT" },
  { name: "RelationshipId", ddl: "RelationshipId TEXT" },
  { name: "ContentHash", ddl: "ContentHash TEXT DEFAULT ''" },
  { name: "Version", ddl: "Version INTEGER DEFAULT 1" },
  { name: "PdfPayload", ddl: "PdfPayload BLOB" },
  { name: "UpdatedUtc", ddl: "UpdatedUtc TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))" },
];

const COLUMN_BACKFILLS: ColumnBackfill[] = [
  { sql: `UPDATE DataFiles SET DocumentID = lower(hex(randomblob(16))) WHERE DocumentID IS NULL OR length(trim(DocumentID)) = 0;` },
  { sql: `UPDATE DataFiles SET ContentType = 'application/pdf' WHERE ContentType IS NULL OR length(trim(ContentType)) = 0;` },
  { sql: `UPDATE DataFiles SET ContentHash = '' WHERE ContentHash IS NULL;` },
  { sql: `UPDATE DataFiles SET Version = 1 WHERE Version IS NULL OR Version < 1;` },
  { sql: `UPDATE DataFiles SET UpdatedUtc = COALESCE(NULLIF(trim(UpdatedUtc), ''), CreatedUtc, strftime('%Y-%m-%dT%H:%M:%fZ','now')) WHERE UpdatedUtc IS NULL OR length(trim(UpdatedUtc)) = 0;` },
];

const DATA_FILES_SELECT = `
  SELECT DocumentID, FileName, ContentType, PartUri, RelationshipId, ContentHash, Version, RawFileSize,
         ThumbnailPng, ThumbnailWidth, ThumbnailHeight, ThumbnailMimeType, CreatedUtc, UpdatedUtc
  FROM DataFiles
  WHERE PartUri IS NOT NULL AND length(trim(PartUri)) > 0
  ORDER BY datetime(CreatedUtc) DESC;`;

const DATA_FILES_UPSERT = `
  INSERT OR REPLACE INTO DataFiles
    (DocumentID, FileName, ContentType, PartUri, RelationshipId, ContentHash, Version, PdfPayload, RawFileSize, ThumbnailPng, ThumbnailWidth, ThumbnailHeight, ThumbnailMimeType, CreatedUtc, UpdatedUtc)
  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);`;

@Injectable({ providedIn: "root" })
export class DataFileRepository {
  constructor(private logger: NGXLogger) {}

  ensureSchema(database: Database): void {
    this.createTableIfNeeded(database);
    this.migrateColumns(database);
    this.createIndexes(database);
  }

  upsert(database: Database, record: DataFileRecord): void {
    database.run(DATA_FILES_UPSERT, [
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
    ]);
  }

  upsertAll(database: Database, records: DataFileRecord[]): void {
    if (!records.length) return;

    try {
      database.run("BEGIN TRANSACTION;");
      for (const record of records) {
        this.upsert(database, record);
      }
      database.run("COMMIT;");
    } catch (error) {
      this.tryRollback(database);
      throw error;
    }
  }

  queryAll(database: Database): DataFileRecord[] {
    try {
      const result = database.exec(DATA_FILES_SELECT);
      const rows = result?.[0]?.values ?? [];
      return rows.map((row) => this.mapRow(row));
    } catch (error) {
      this.logger.error("queryAll: failed to query DataFiles", error);
      return [];
    }
  }

  deleteByPartUri(database: Database, partUri: string): void {
    database.run(`DELETE FROM DataFiles WHERE PartUri = ?;`, [partUri]);
  }

  findRelationshipId(database: Database, partUri: string): string | null {
    const lookup = database.exec(`SELECT RelationshipId FROM DataFiles WHERE PartUri = ? LIMIT 1;`, [partUri]);
    const value = lookup?.[0]?.values?.[0]?.[0];
    return value !== null && value !== undefined ? String(value) : null;
  }

  loadPayloadByPartUri(database: Database, partUri: string): Uint8Array | null {
    try {
      const row = database.exec(`SELECT PdfPayload FROM DataFiles WHERE PartUri = ? LIMIT 1;`, [partUri]);
      const payload = row?.[0]?.values?.[0]?.[0];
      if (payload instanceof Uint8Array && payload.length > 0) {
        return new Uint8Array(payload);
      }
    } catch (error) {
      this.logger.warn("loadPayloadByPartUri: failed to read PdfPayload blob", error);
    }
    return null;
  }

  private mapRow(row: unknown[]): DataFileRecord {
    const mime = row[11];
    const thumbnailMimeType = typeof mime === "string" && mime.trim() ? mime : "image/png";
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
  }

  private createTableIfNeeded(database: Database): void {
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
  }

  private migrateColumns(database: Database): void {
    const existing = this.getExistingColumns(database);
    this.addMissingColumns(database, existing);
    this.runBackfills(database);
  }

  private getExistingColumns(database: Database): Set<string> {
    const info = database.exec("PRAGMA table_info(DataFiles);");
    return new Set((info?.[0]?.values ?? []).map((row) => String(row[1]).toLowerCase()));
  }

  private addMissingColumns(database: Database, existing: Set<string>): void {
    for (const migration of COLUMN_MIGRATIONS) {
      if (!existing.has(migration.name.toLowerCase())) {
        database.run(`ALTER TABLE DataFiles ADD COLUMN ${migration.ddl};`);
      }
    }
  }

  private runBackfills(database: Database): void {
    for (const backfill of COLUMN_BACKFILLS) {
      database.run(backfill.sql);
    }
  }

  private createIndexes(database: Database): void {
    database.run(`CREATE INDEX IF NOT EXISTS IX_DataFiles_PartUri ON DataFiles(PartUri);`);
    database.run(`CREATE INDEX IF NOT EXISTS IX_DataFiles_RelationshipId ON DataFiles(RelationshipId);`);
    database.run(`CREATE INDEX IF NOT EXISTS IX_DataFiles_DocumentId ON DataFiles(DocumentID);`);
  }

  private tryRollback(database: Database): void {
    try {
      database.run("ROLLBACK;");
    } catch (rollbackError) {
      this.logger.warn("tryRollback: rollback failed", rollbackError);
    }
  }
}
