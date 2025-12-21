/* global Office, Excel */

import { Component, NgZone } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import { DataService } from "./services/data.service";
import { sql } from "@codemirror/lang-sql";
import brotliModulePromise from "brotli-wasm";
import { GlobalWorkerOptions, getDocument } from "pdfjs-dist";
const pdfWorkerSrc = new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).toString();
import seedScript from "../seed-script.sql";

type UploadStatus = "Queued" | "In Progress" | "Complete";
interface UploadItem {
  file?: File;
  fileName: string;
  status: UploadStatus;
  progress: number;
  rawFileSize?: number;
  compressedFileSize?: number;
  xmlPartName?: string;
  createdUtc?: string;
  isDeleting?: boolean;
}

interface PdfThumbnail {
  fileName: string;
  xmlPartName: string;
  createdUtc?: string;
  imageUrl: string;
}

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"],
})
export class AppComponent {
  isReady = false;
  isExcelHost = false;
  isRunning = false;
  isCreatingDb = false;
  statusMessage = "";
  lastAddress = "";
  activeTab: "sql" | "pdf" | "pdfViewer" = "sql";
  hasData = false;
  readonly requiredTables = ["Pages", "Cells", "PolygonData"];
  missingRequiredTables: string[] = [...this.requiredTables];
  private brotliInstance?: any;
  private readonly seedDisplayQuery = `SELECT
    p.PageKey,
    p.PdfPageNumber,
    c.CellID,
    c.CellName,
    c.Value,
    json_group_array(
        json_object(
            'X', pd.X,
            'Y', pd.Y
        )
        ORDER BY pd.PointOrder
  ) AS Polygon
FROM Pages p
JOIN Cells c
    ON c.PageID = p.PageID
JOIN PolygonData pd
    ON pd.CellPK = c.CellPK
WHERE p.PageKey = '1.1'
GROUP BY
    c.CellPK
ORDER BY
    c.ExcelRow,
    c.ExcelColumn;`;
  private readonly seedPlaceholder = "Please seed the Sqlite database";
  sqlInput = this.seedDisplayQuery;
  sqlExtensions = [sql({ upperCaseKeywords: true })];
  queryResults: Array<{ columns: string[]; values: unknown[][] }> = [];
  hasDatabase = false;
  isSeeding = false;
  // tracks if file uploads are being processed sequentially
  isUploading = false;
  hasDataFiles = false;
  isLoadingThumbnails = false;
  readonly uploadSteps = 6;
  uploads: UploadItem[] = [];
  pdfThumbnails: PdfThumbnail[] = [];

  get canSeedDatabase(): boolean {
    const missingAllRequiredTables = this.requiredTables.every((table) =>
      this.missingRequiredTables.includes(table),
    );
    return !this.hasDatabase || missingAllRequiredTables;
  }

  constructor(
    private zone: NgZone,
    private logger: NGXLogger,
    private dataService: DataService,
  ) {
    GlobalWorkerOptions.workerSrc = pdfWorkerSrc;
    if (typeof Office !== "undefined" && Office.onReady) {
      Office.onReady((info) => {
        this.zone.run(() => {
          this.isExcelHost = info.host === Office.HostType.Excel;
          this.isReady = true;
          this.logger.debug("Office is ready");
          this.checkDatabaseState();
        });
      });
    } else {
      this.zone.run(() => {
        this.isReady = true;
      });
    }
  }

  async run(): Promise<void> {
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to run SQL.";
      return;
    }

    if (!this.hasDatabase) {
      this.statusMessage = "Create or seed the database before running SQL.";
      return;
    }

    if (!this.hasData) {
      this.statusMessage = "Seed the database before running SQL.";
      return;
    }

    this.isRunning = true;
    this.statusMessage = "Running SQL...";

    try {
      const results = await this.dataService.execute(this.sqlInput);
      this.zone.run(() => {
        this.queryResults = results;
        this.statusMessage = results.length ? "Query completed." : "Query returned no results.";
        this.hasDatabase = true;
        if (!this.hasData && results.some((r) => r.values?.length)) {
          this.hasData = true;
        }
      });
    } catch (error) {
      console.error(error);
      this.zone.run(() => {
        this.statusMessage = `Error: ${error instanceof Error ? error.message : String(error)}`;
        this.queryResults = [];
      });
    } finally {
      this.zone.run(() => {
        this.isRunning = false;
      });
    }
  }

  async seedDatabase(): Promise<void> {
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to seed the database.";
      return;
    }

    this.isSeeding = true;
    this.statusMessage = "Seeding database...";

    try {
      await this.dataService.seedDatabase(seedScript);
      this.zone.run(() => {
        this.hasDatabase = true;
        this.hasData = true;
        this.missingRequiredTables = [];
        this.sqlInput = this.seedDisplayQuery;
        this.statusMessage = "Database seeded.";
      });
    } catch (error) {
      console.error(error);
      this.zone.run(() => {
        this.statusMessage = `Seed error: ${error instanceof Error ? error.message : String(error)}`;
      });
    } finally {
      this.zone.run(() => {
        this.isSeeding = false;
      });
    }
  }

  private async checkDatabaseState(): Promise<void> {
    if (!this.isExcelHost) {
      this.hasDatabase = false;
      this.hasData = false;
      this.missingRequiredTables = [...this.requiredTables];
      return;
    }

    try {
      const state = await this.dataService.getDatabaseState(this.requiredTables);
      this.zone.run(() => {
        this.hasDatabase = state.hasDatabase;
        this.hasData = state.hasData;
        this.missingRequiredTables = state.missingRequiredTables ?? [];
        if (!state.hasDatabase || !state.hasData) {
          this.sqlInput = this.seedPlaceholder;
        } else if (!this.sqlInput.trim() || this.sqlInput === this.seedPlaceholder) {
          this.sqlInput = this.seedDisplayQuery;
        }
      });
    } catch (error) {
      this.logger.error("Failed to check database state", error);
      this.zone.run(() => {
        this.hasDatabase = false;
        this.hasData = false;
        this.missingRequiredTables = [...this.requiredTables];
      });
    }
  }

  async createDatabase(): Promise<void> {
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to create the database.";
      return;
    }

    this.isCreatingDb = true;
    this.statusMessage = "Creating database...";

    try {
      await this.dataService.loadOrCreate();
      this.statusMessage = "Database is ready.";
      this.logger.info("Database created or loaded.");
    } catch (error) {
      this.logger.error("Failed to create database", error);
      this.statusMessage = `Database error: ${error instanceof Error ? error.message : String(error)}`;
    } finally {
      this.isCreatingDb = false;
    }
  }

  openPdfPicker(fileInput: HTMLInputElement): void {
    fileInput.value = "";
    fileInput.click();
  }

  onPdfSelected(event: Event): void {
    const input = event.target as HTMLInputElement | null;
    const files = input?.files ? Array.from(input.files) : [];
    this.uploads = files.map<UploadItem>((file) => ({
      file,
      fileName: file.name,
      status: "Queued",
      progress: 0,
    }));

    if (!files.length) {
      return;
    }

    if (!this.isReady) {
      this.statusMessage = "Wait for Office to finish loading before uploading files.";
      return;
    }

    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to upload files.";
      return;
    }

    void this.processUploadsSequentially();
  }

  private async processUploadsSequentially(): Promise<void> {
    if (this.isUploading) {
      return;
    }

    this.isUploading = true;
    for (const upload of this.uploads) {
      try {
        await this.processSingleUpload(upload);
      } catch (error) {
        this.logger.error("File upload failed", error);
        this.zone.run(() => {
          upload.status = "Queued";
          upload.progress = 0;
          this.uploads = [...this.uploads];
          this.statusMessage = `Upload failed for ${upload.fileName}: ${
            error instanceof Error ? error.message : String(error)
          }`;
        });
      }
    }

    this.zone.run(() => {
      this.isUploading = false;
    });
  }

  private async processSingleUpload(upload: UploadItem): Promise<void> {
    const totalSteps = this.uploadSteps;
    this.updateUploadProgress(upload, 0, totalSteps, "In Progress");

    // 1. Get the file name and file size
    if (!upload.file) {
      throw new Error("No file content available for upload.");
    }

    const rawFileSize = upload.file.size;
    upload.rawFileSize = rawFileSize;
    this.updateUploadProgress(upload, 1, totalSteps);

    // 2. Compress the file using brotli
    const buffer = await upload.file.arrayBuffer();
    const compressedBytes = await this.compressWithBrotli(new Uint8Array(buffer));
    this.updateUploadProgress(upload, 2, totalSteps);

    // 3. Stringify the file
    const base64Payload = this.bytesToBase64(compressedBytes);
    this.updateUploadProgress(upload, 3, totalSteps);

    // 4. Get the compressed file size
    const compressedFileSize = compressedBytes.byteLength;
    upload.compressedFileSize = compressedFileSize;
    this.updateUploadProgress(upload, 4, totalSteps);

    // 5. Save the stringified file as a customXml part in the Excel document
    const { xmlPartName, createdUtc } = await this.dataService.saveFilePart(upload.fileName, base64Payload);
    upload.xmlPartName = xmlPartName;
    upload.createdUtc = createdUtc;
    this.updateUploadProgress(upload, 5, totalSteps);

    // 6. Create an entry in the DataFiles table
    await this.dataService.recordDataFile({
      fileName: upload.fileName,
      xmlPartName,
      rawFileSize,
      compressedFileSize,
      createdUtc,
    });

    this.updateUploadProgress(upload, totalSteps, totalSteps, "Complete");
    this.zone.run(() => {
      this.hasDatabase = true;
      this.hasDataFiles = true;
      this.statusMessage = "File upload recorded.";
    });
  }

  private updateUploadProgress(upload: UploadItem, step: number, totalSteps: number, status?: UploadStatus): void {
    const progress = Math.min(100, Math.round((step / totalSteps) * 100));
    this.zone.run(() => {
      if (status) {
        upload.status = status;
      }
      upload.progress = progress;
      this.uploads = [...this.uploads];
    });
  }

  private bytesToBase64(bytes: Uint8Array): string {
    const chunkSize = 0x8000;
    let binary = "";
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode(...chunk);
    }
    return btoa(binary);
  }

  private base64ToBytes(base64: string): Uint8Array {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  private async renderPdfThumbnail(pdfBytes: Uint8Array): Promise<string> {
    const loadingTask = getDocument({ data: pdfBytes });
    const pdf = await loadingTask.promise;
    const page = await pdf.getPage(1);
    const baseViewport = page.getViewport({ scale: 1 });
    const targetWidth = 260;
    const scale = Math.min(1.5, targetWidth / baseViewport.width);
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement("canvas");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const context = canvas.getContext("2d");
    if (!context) {
      throw new Error("Unable to render PDF thumbnail.");
    }

    await page.render({ canvasContext: context, viewport, canvas }).promise;
    const dataUrl = canvas.toDataURL("image/png");
    pdf.destroy();
    return dataUrl;
  }

  private async compressWithBrotli(input: Uint8Array): Promise<Uint8Array> {
    const brotli = await this.ensureBrotli();
    return brotli.compress(input);
  }

  private async decompressWithBrotli(input: Uint8Array): Promise<Uint8Array> {
    const brotli = await this.ensureBrotli();
    return brotli.decompress(input);
  }

  private async ensureBrotli(): Promise<any> {
    if (!this.brotliInstance) {
      this.brotliInstance = await (brotliModulePromise as unknown as Promise<any>);
    }
    return this.brotliInstance;
  }

  async deleteUpload(upload: UploadItem): Promise<void> {
    if (!this.isReady) {
      this.statusMessage = "Wait for Office to finish loading before deleting files.";
      return;
    }
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to delete files.";
      return;
    }
    if (!upload.xmlPartName) {
      this.statusMessage = "File has not been saved yet; nothing to delete.";
      return;
    }

    this.zone.run(() => {
      upload.isDeleting = true;
      upload.status = "In Progress";
      upload.progress = 0;
      this.uploads = [...this.uploads];
    });

    try {
      await this.dataService.deleteDataFile(upload.xmlPartName);
      this.zone.run(() => {
        this.uploads = this.uploads.filter((item) => item !== upload);
        this.pdfThumbnails = this.pdfThumbnails.filter((thumb) => thumb.xmlPartName !== upload.xmlPartName);
        this.hasDataFiles =
          this.uploads.some((item) => Boolean(item.xmlPartName)) || this.pdfThumbnails.length > 0;
        this.statusMessage = `Deleted ${upload.fileName}.`;
      });
    } catch (error) {
      this.logger.error("Failed to delete upload", error);
      this.zone.run(() => {
        upload.isDeleting = false;
        upload.status = "Complete";
        upload.progress = 100;
        this.uploads = [...this.uploads];
        this.statusMessage = `Delete failed for ${upload.fileName}: ${
          error instanceof Error ? error.message : String(error)
        }`;
      });
    }
  }

  setTab(tab: "sql" | "pdf" | "pdfViewer"): void {
    if (tab === "pdfViewer" && !this.hasDataFiles) {
      this.statusMessage = "Upload a PDF first to enable the viewer.";
      return;
    }

    this.activeTab = tab;
    if (tab === "pdf" || tab === "pdfViewer") {
      void this.syncUploadsFromDatabase();
    }
    if (tab === "pdfViewer") {
      void this.loadPdfThumbnails();
    }
  }

  private async syncUploadsFromDatabase(): Promise<void> {
    if (!this.isReady || !this.isExcelHost) {
      return;
    }

    try {
      const records = await this.dataService.getDataFiles();
      this.zone.run(() => {
        const existing = [...this.uploads];
        const updated = new Map<string, UploadItem>();

        // Seed with current uploads for quick lookup
        for (const item of existing) {
          if (item.xmlPartName) {
            updated.set(item.xmlPartName, item);
          }
        }

        for (const record of records) {
          const existingItem = updated.get(record.xmlPartName);
          if (existingItem) {
            existingItem.status = "Complete";
            existingItem.progress = 100;
            existingItem.fileName = record.fileName;
            existingItem.rawFileSize = record.rawFileSize;
            existingItem.compressedFileSize = record.compressedFileSize;
            existingItem.createdUtc = record.createdUtc;
          } else {
            updated.set(record.xmlPartName, {
              fileName: record.fileName,
              status: "Complete",
              progress: 100,
              rawFileSize: record.rawFileSize,
              compressedFileSize: record.compressedFileSize,
              xmlPartName: record.xmlPartName,
              createdUtc: record.createdUtc,
            });
          }
        }

        // Keep any in-flight uploads (no xmlPartName yet)
        const inflight = existing.filter((item) => !item.xmlPartName);
        this.uploads = [...updated.values(), ...inflight];
        this.hasDataFiles = [...updated.values()].length > 0;
      });
    } catch (error) {
      this.logger.error("Failed to sync uploads from database", error);
      this.zone.run(() => {
        this.statusMessage = `Unable to load uploaded files: ${error instanceof Error ? error.message : String(error)}`;
      });
    }
  }

  private async loadPdfThumbnails(): Promise<void> {
    if (!this.isReady || !this.isExcelHost) {
      return;
    }

    this.zone.run(() => {
      this.isLoadingThumbnails = true;
    });

    try {
      const records = await this.dataService.getDataFiles();
      const thumbnails: PdfThumbnail[] = [];
      this.zone.run(() => {
        this.hasDataFiles = records.length > 0;
      });

      for (const record of records) {
        const payloadBase64 = await this.dataService.loadFilePart(record.xmlPartName);
        if (!payloadBase64) {
          continue;
        }
        const compressedBytes = this.base64ToBytes(payloadBase64);
        const pdfBytes = await this.decompressWithBrotli(compressedBytes);
        const imageUrl = await this.renderPdfThumbnail(pdfBytes);
        thumbnails.push({
          fileName: record.fileName,
          xmlPartName: record.xmlPartName,
          createdUtc: record.createdUtc,
          imageUrl,
        });
      }

      this.zone.run(() => {
        this.pdfThumbnails = thumbnails;
        this.hasDataFiles = this.hasDataFiles || thumbnails.length > 0;
      });
    } catch (error) {
      this.logger.error("Failed to load PDF thumbnails", error);
      this.zone.run(() => {
        this.statusMessage = `Unable to load PDF thumbnails: ${error instanceof Error ? error.message : String(error)}`;
      });
    } finally {
      this.zone.run(() => {
        this.isLoadingThumbnails = false;
      });
    }
  }
}
