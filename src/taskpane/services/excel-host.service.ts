import { Injectable } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import { normalizeSliceData, concatUint8Arrays, uint8ArrayToBase64, base64ToUint8Array } from "./utils/binary-encoding";
import { extractBase64FromXml } from "./utils/xml-helpers";
import initSqlJs, { Database, SqlJsStatic } from "sql.js";

/* global Excel, Office */

const CUSTOM_XML_ELEMENT = "proofPanelData";

@Injectable({ providedIn: "root" })
export class ExcelHostService {
  private officeReadyPromise?: Promise<unknown>;
  private sqlPromise?: Promise<SqlJsStatic>;

  constructor(private logger: NGXLogger) {}

  async waitForOfficeReady(): Promise<boolean> {
    if (!this.officeReadyPromise) {
      if (typeof Office === "undefined" || !Office.onReady) return false;
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

  async ensureExcelReady(): Promise<void> {
    const ready = await this.waitForOfficeReady();
    if (!ready || typeof Excel === "undefined") {
      throw new Error("Excel is not available. Connect to Excel before accessing the database.");
    }
  }

  async findDataPart(context: Excel.RequestContext): Promise<Excel.CustomXmlPart | null> {
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

  async saveCustomXmlPayload(database: Database): Promise<void> {
    this.logger.info("saveCustomXmlPayload: exporting database and writing to customXml");
    const payload = uint8ArrayToBase64(database.export());
    const xml = `<?xml version="1.0" encoding="UTF-8"?><${CUSTOM_XML_ELEMENT}>${payload}</${CUSTOM_XML_ELEMENT}>`;

    await Excel.run(async (context) => {
      const existingPart = await this.findDataPart(context);
      if (existingPart) {
        existingPart.delete();
      }
      context.workbook.customXmlParts.add(xml);
      await context.sync();
    });
    this.logger.info("saveCustomXmlPayload: database saved to customXml");
  }

  async deleteCustomXmlPayload(): Promise<void> {
    await Excel.run(async (context) => {
      const existingPart = await this.findDataPart(context);
      if (existingPart) {
        existingPart.delete();
        await context.sync();
      }
    });
  }

  async readCustomXmlPayload(): Promise<string | null> {
    let xmlPayload: string | null = null;

    await Excel.run(async (context) => {
      const dataPart = await this.findDataPart(context);
      if (!dataPart) {
        this.logger.debug("readCustomXmlPayload: no data part found");
        return;
      }
      const xmlResult = dataPart.getXml();
      await context.sync();
      xmlPayload = xmlResult.value ?? null;
    });

    return xmlPayload;
  }

  extractDatabaseBytes(xmlPayload: string): Uint8Array | null {
    const base64 = extractBase64FromXml(xmlPayload, CUSTOM_XML_ELEMENT);
    if (!base64) return null;
    return base64ToUint8Array(base64);
  }

  async loadSqlEngine(): Promise<SqlJsStatic> {
    this.logger.debug("loadSqlEngine: loading sql.js");
    if (!this.sqlPromise) {
      this.logger.debug("loadSqlEngine: initializing sql.js");
      this.sqlPromise = initSqlJs({ locateFile: (file) => file });
    }
    return this.sqlPromise;
  }

  async readCompressedWorkbookBytes(): Promise<Uint8Array> {
    const file = await this.getCompressedFileHandle();
    try {
      const slices = await this.readAllSlices(file);
      return concatUint8Arrays(slices);
    } finally {
      await this.closeFileHandle(file);
    }
  }

  private async readAllSlices(file: Office.File): Promise<Uint8Array[]> {
    const slices: Uint8Array[] = [];
    for (let index = 0; index < file.sliceCount; index += 1) {
      const slice = await this.getFileSlice(file, index);
      slices.push(normalizeSliceData(slice.data));
    }
    return slices;
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
          reject(new Error(result.error?.message ?? "Unable to read workbook package from the Office host."));
        },
      );
    });
  }

  private getFileSlice(file: Office.File, index: number): Promise<Office.Slice> {
    return new Promise((resolve, reject) => {
      file.getSliceAsync(index, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
          return;
        }
        reject(new Error(result.error?.message ?? `Unable to read workbook package slice at index ${index}.`));
      });
    });
  }

  private closeFileHandle(file: Office.File): Promise<void> {
    return new Promise((resolve) => {
      file.closeAsync(() => resolve());
    });
  }
}
