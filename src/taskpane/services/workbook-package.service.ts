import { Injectable } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { uint8ArrayToBase64 } from "./utils/binary-encoding";
import { parseXmlDocument, serializeXmlDocument } from "./utils/xml-helpers";
import { ExcelHostService } from "./excel-host.service";

/* global Excel */

const PACKAGE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const CONTENT_TYPES_PATH = "[Content_Types].xml";
const RELS_PATH = "xl/_rels/workbook.xml.rels";
const EMPTY_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

type PackageEntries = Record<string, Uint8Array>;

@Injectable({ providedIn: "root" })
export class WorkbookPackageService {
  private pendingEntries?: PackageEntries;
  private isPendingDirty = false;

  constructor(
    private logger: NGXLogger,
    private excelHost: ExcelHostService,
  ) {}

  normalizePartUri(partUri: string): string {
    const trimmed = String(partUri ?? "").trim().replace(/\\/g, "/");
    if (!trimmed) {
      throw new Error("Part URI cannot be empty.");
    }
    return trimmed.startsWith("/") ? trimmed : `/${trimmed}`;
  }

  partUriToZipPath(partUri: string): string {
    return this.normalizePartUri(partUri).replace(/^\//, "");
  }

  hasPendingBatch(): boolean {
    return Boolean(this.pendingEntries);
  }

  async beginBatch(): Promise<void> {
    if (this.pendingEntries) return;
    this.pendingEntries = await this.readPackage();
    this.isPendingDirty = false;
    this.logger.info("beginBatch: initialized staged workbook package");
  }

  async commitBatch(): Promise<void> {
    if (!this.pendingEntries) return;

    if (this.isPendingDirty) {
      await this.writePackage(this.pendingEntries);
      this.logger.info("commitBatch: committed staged workbook package");
    } else {
      this.logger.info("commitBatch: no staged workbook changes to commit");
    }
    this.resetBatch();
  }

  rollbackBatch(): void {
    this.resetBatch();
    this.logger.info("rollbackBatch: cleared staged workbook package");
  }

  markPendingDirty(): void {
    this.isPendingDirty = true;
  }

  async getEntriesForWrite(deferWrite: boolean): Promise<PackageEntries> {
    if (deferWrite) {
      if (!this.pendingEntries) {
        this.pendingEntries = await this.readPackage();
        this.isPendingDirty = false;
      }
      return this.pendingEntries;
    }
    return this.readPackage();
  }

  async writePackage(entries: PackageEntries): Promise<void> {
    const compressed = zipSync(entries, { level: 6 });
    await Excel.createWorkbook(uint8ArrayToBase64(compressed));
  }

  async readPackage(): Promise<PackageEntries> {
    const compressed = await this.excelHost.readCompressedWorkbookBytes();
    return unzipSync(compressed);
  }

  addOrUpdateContentTypeOverride(entries: PackageEntries, partUri: string, contentType: string): void {
    const xml = entries[CONTENT_TYPES_PATH];
    if (!xml) {
      throw new Error("Workbook package is missing [Content_Types].xml.");
    }

    const doc = parseXmlDocument(strFromU8(xml));
    const normalizedUri = this.normalizePartUri(partUri);
    const override = this.findOrCreateElement(doc, "Override", "PartName", normalizedUri);
    override.setAttribute("PartName", normalizedUri);
    override.setAttribute("ContentType", contentType);
    entries[CONTENT_TYPES_PATH] = strToU8(serializeXmlDocument(doc));
  }

  removeContentTypeOverride(entries: PackageEntries, partUri: string): void {
    const xml = entries[CONTENT_TYPES_PATH];
    if (!xml) return;

    const doc = parseXmlDocument(strFromU8(xml));
    const normalizedUri = this.normalizePartUri(partUri);
    this.removeElementsByAttribute(doc, "Override", "PartName", normalizedUri);
    entries[CONTENT_TYPES_PATH] = strToU8(serializeXmlDocument(doc));
  }

  addOrUpdateRelationship(entries: PackageEntries, partUri: string, relationshipId: string): void {
    const xml = entries[RELS_PATH] ? strFromU8(entries[RELS_PATH]) : EMPTY_RELS_XML;
    const doc = parseXmlDocument(xml);
    const target = this.partUriToRelTarget(partUri);
    const relationship = this.findOrCreateElement(doc, "Relationship", "Id", relationshipId);
    relationship.setAttribute("Id", relationshipId);
    relationship.setAttribute("Type", PACKAGE_RELATIONSHIP_TYPE);
    relationship.setAttribute("Target", target);
    entries[RELS_PATH] = strToU8(serializeXmlDocument(doc));
  }

  removeRelationship(entries: PackageEntries, partUri: string, relationshipId?: string | null): void {
    const xml = entries[RELS_PATH];
    if (!xml) return;

    const doc = parseXmlDocument(strFromU8(xml));
    const target = this.partUriToRelTarget(partUri);
    const relationships = Array.from(doc.getElementsByTagName("Relationship"));
    for (const rel of relationships) {
      const matchesId = Boolean(relationshipId) && rel.getAttribute("Id") === relationshipId;
      const matchesTarget = rel.getAttribute("Target") === target;
      if (matchesId || matchesTarget) {
        rel.parentNode?.removeChild(rel);
      }
    }
    entries[RELS_PATH] = strToU8(serializeXmlDocument(doc));
  }

  async removePackagePart(
    entries: PackageEntries,
    partUri: string,
    relationshipId: string | null,
  ): Promise<void> {
    delete entries[this.partUriToZipPath(partUri)];
    this.removeContentTypeOverride(entries, partUri);
    this.removeRelationship(entries, partUri, relationshipId);
    await this.writePackage(entries);
  }

  private partUriToRelTarget(partUri: string): string {
    return `../${this.partUriToZipPath(partUri)}`;
  }

  private findOrCreateElement(doc: Document, tagName: string, matchAttr: string, matchValue: string): Element {
    const elements = Array.from(doc.getElementsByTagName(tagName));
    const existing = elements.find((el) => el.getAttribute(matchAttr) === matchValue);
    if (existing) return existing;

    const created = doc.createElementNS(doc.documentElement.namespaceURI, tagName);
    doc.documentElement.appendChild(created);
    return created;
  }

  private removeElementsByAttribute(doc: Document, tagName: string, attrName: string, attrValue: string): void {
    const elements = Array.from(doc.getElementsByTagName(tagName));
    for (const el of elements) {
      if (el.getAttribute(attrName) === attrValue) {
        el.parentNode?.removeChild(el);
      }
    }
  }

  private resetBatch(): void {
    this.pendingEntries = undefined;
    this.isPendingDirty = false;
  }
}
