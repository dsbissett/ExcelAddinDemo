const PDF_PARTS_FOLDER = "pdfs";

export function createDocumentId(): string {
  const cryptoApi = globalThis.crypto;
  if (cryptoApi?.randomUUID) {
    return cryptoApi.randomUUID();
  }
  return `doc-${Date.now().toString(36)}-${randomToken(12)}`;
}

export function createPdfPartUri(fileName: string): string {
  const stripped = fileName.replace(/\.pdf$/i, "");
  const safeName =
    stripped
      .replace(/[^a-zA-Z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase() || "document";
  return `/${PDF_PARTS_FOLDER}/${safeName}-${randomToken(8)}.pdf`;
}

export function createRelationshipId(): string {
  return `rIdPdf${randomToken(10)}`;
}

export function randomToken(length: number): string {
  const value = Math.random().toString(36).slice(2);
  if (value.length >= length) {
    return value.slice(0, length);
  }
  return `${value}${Math.random().toString(36).slice(2)}`.slice(0, length);
}
