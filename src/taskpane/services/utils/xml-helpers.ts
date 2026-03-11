export function parseXmlDocument(xmlText: string): Document {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, "application/xml");
  if (doc.querySelector("parsererror")) {
    throw new Error("Unable to parse workbook package XML.");
  }
  return doc;
}

export function serializeXmlDocument(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

export function extractBase64FromXml(xml: string, elementName: string): string | null {
  return tryExtractViaDom(xml, elementName) ?? tryExtractViaString(xml, elementName);
}

function tryExtractViaDom(xml: string, elementName: string): string | null {
  try {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.querySelector("parsererror")) return null;
    const byName = doc.querySelector(elementName);
    const payload = (byName ?? doc.documentElement)?.textContent?.trim() ?? "";
    return payload || null;
  } catch {
    return null;
  }
}

function tryExtractViaString(xml: string, elementName: string): string | null {
  const openTag = `<${elementName}`;
  const start = xml.indexOf(openTag);
  if (start === -1) return null;

  const startContent = xml.indexOf(">", start);
  const end = xml.indexOf(`</${elementName}>`, startContent);
  if (startContent === -1 || end === -1 || end <= startContent) return null;

  const payload = xml.slice(startContent + 1, end).trim();
  return payload || null;
}
