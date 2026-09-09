import { strFromU8, strToU8 } from "fflate";
import { preserveNamedWorkbookStyles } from "./preserved-styles";

type ArchiveEntries = Record<string, Uint8Array>;
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
const XMLNS = "http://www.w3.org/2000/xmlns/";

function fail(reason: string): never { throw new Error(`xlsx_cell_layout_preservation_failed:${reason}`); }
function elements(node: Node): Element[] { return Array.from(node.childNodes).filter((item): item is Element => item.nodeType === 1); }
function child(node: Node, name: string): Element | undefined { return elements(node).find((item) => item.namespaceURI === NS && item.localName === name); }
function parse(bytes: Uint8Array | undefined, part: string): Document {
  if (!bytes) fail(`missing_${part}`);
  const xml = strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) fail("unsafe_xml");
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length || !document.documentElement) fail(`invalid_${part}`);
  return document;
}
function encode(document: Document): Uint8Array { return strToU8(new XMLSerializer().serializeToString(document)); }
function signature(node: Node): string {
  if (node.nodeType === 3 || node.nodeType === 4) return node.nodeValue?.trim() ? JSON.stringify(node.nodeValue) : "";
  if (node.nodeType !== 1) return "";
  const element = node as Element;
  const attributes = Array.from(element.attributes).filter((item) => item.namespaceURI !== XMLNS)
    .map((item) => [item.namespaceURI, item.localName, item.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
  return JSON.stringify([element.namespaceURI, element.localName, attributes, Array.from(element.childNodes).map(signature).filter(Boolean)]);
}
function index(value: string | null): number {
  const result = Number(value || 0);
  if (!Number.isSafeInteger(result) || result < 0 || result > 65_535) fail("invalid_style_reference");
  return result;
}
function resolveTarget(target: string | null, base: string): string {
  if (!target || target.includes("\\") || /^[a-z]+:/i.test(target)) fail("invalid_part_target");
  const url = new URL(target, `https://xlsx.invalid/${base}`);
  if (url.origin !== "https://xlsx.invalid" || url.search || url.hash) fail("invalid_part_target");
  return decodeURIComponent(url.pathname.slice(1));
}
function sheets(archive: ArchiveEntries): Array<{ name: string; path: string }> {
  const book = parse(archive["xl/workbook.xml"], "workbook");
  const rels = parse(archive["xl/_rels/workbook.xml.rels"], "workbook_relationships");
  const container = child(book.documentElement, "sheets");
  if (!container) fail("missing_sheets");
  return elements(container).map((sheet) => {
    const id = sheet.getAttributeNS(REL, "id");
    const relationship = elements(rels.documentElement).find((entry) => entry.getAttribute("Id") === id);
    if (!relationship || relationship.getAttribute("Type") !== `${REL}/worksheet` || relationship.getAttribute("TargetMode") === "External") fail("invalid_sheet_relationship");
    return { name: sheet.getAttribute("name") || "", path: resolveTarget(relationship.getAttribute("Target"), "xl/workbook.xml") };
  });
}

/** Retain original row/column styles while generated cells keep their own IDs. */
function preserveStyleReferences(source: ArchiveEntries, target: ArchiveEntries): { bytes: Uint8Array; originalIds: Map<number, number> } {
  const original = parse(source["xl/styles.xml"], "source_styles");
  const merged = parse(strToU8(preserveNamedWorkbookStyles(strFromU8(source["xl/styles.xml"]), strFromU8(target["xl/styles.xml"]))), "merged_styles");
  const componentMaps = new Map<string, Map<number, number>>();
  for (const name of ["fonts", "fills", "borders", "numFmts"]) {
    const from = child(original.documentElement, name), to = child(merged.documentElement, name);
    if (!from) continue;
    if (!to) fail(`missing_style_${name}`);
    const entries = elements(to);
    componentMaps.set(name, new Map(elements(from).map((entry, offset) => {
      const match = entries.findIndex((item) => name === "numFmts" ? item.getAttribute("formatCode") === entry.getAttribute("formatCode") : signature(item) === signature(entry));
      if (match < 0) fail(`unmapped_style_${name}`);
      return [name === "numFmts" ? index(entry.getAttribute("numFmtId")) : offset, name === "numFmts" ? index(entries[match].getAttribute("numFmtId")) : match];
    })));
  }
  const remap = (entry: Element): Element => {
    const copy = merged.importNode(entry, true) as Element;
    for (const [attribute, name] of [["fontId", "fonts"], ["fillId", "fills"], ["borderId", "borders"], ["numFmtId", "numFmts"]]) {
      const id = index(copy.getAttribute(attribute));
      if (name === "numFmts" && id < 164) continue;
      const mapped = componentMaps.get(name)?.get(id);
      if (mapped === undefined) fail(`unmapped_${attribute}`);
      copy.setAttribute(attribute, String(mapped));
    }
    return copy;
  };
  const originalNamed = child(original.documentElement, "cellStyleXfs"), targetNamed = child(merged.documentElement, "cellStyleXfs");
  const originalCells = child(original.documentElement, "cellXfs"), targetCells = child(merged.documentElement, "cellXfs");
  if (!originalNamed || !targetNamed || !originalCells || !targetCells) fail("missing_cell_styles");
  const targetNamedEntries = elements(targetNamed);
  const namedIds = new Map(elements(originalNamed).map((entry, offset) => {
    const key = signature(remap(entry));
    const match = targetNamedEntries.findIndex((item) => signature(item) === key);
    if (match < 0) fail("unmapped_named_style");
    return [offset, match];
  }));
  const originalIds = new Map<number, number>();
  const targetEntries = elements(targetCells);
  for (const [offset, entry] of elements(originalCells).entries()) {
    const copy = remap(entry);
    const named = namedIds.get(index(copy.getAttribute("xfId")));
    if (named === undefined) fail("unmapped_cell_parent");
    copy.setAttribute("xfId", String(named));
    if (offset === 0) {
      // Duke reserves cell XF 0 for the default style. Explicit user formats
      // receive their own XFs; preserving this slot retains the imported theme.
      if (!targetEntries[0]) fail("missing_default_cell_style");
      targetCells.replaceChild(copy, targetEntries[0]);
      targetEntries[0] = copy;
      originalIds.set(0, 0);
      continue;
    }
    const key = signature(copy);
    const match = targetEntries.findIndex((item) => signature(item) === key);
    if (match >= 0) originalIds.set(offset, match);
    else {
      if (targetEntries.length >= 65_490) fail("cell_style_limit");
      originalIds.set(offset, targetEntries.length);
      targetCells.appendChild(copy);
      targetEntries.push(copy);
    }
  }
  targetCells.setAttribute("count", String(targetEntries.length));
  // Source conditional/table styles reference the source differential pool.
  // These are unchanged when the caller proves a cell-only mutation history.
  for (const name of ["dxfs", "tableStyles", "colors", "extLst"]) {
    const existing = child(merged.documentElement, name);
    if (existing) merged.documentElement.removeChild(existing);
    const preserved = child(original.documentElement, name);
    if (preserved) merged.documentElement.appendChild(merged.importNode(preserved, true));
  }
  return { bytes: encode(merged), originalIds };
}

/**
 * Called only while the controller's structural revision equals its hydrated
 * source baseline. Cell values/formats may change; layout and all other OOXML
 * parts remain the source's. The independent worksheet-name/order check fails
 * closed if a caller accidentally supplies a different workbook.
 */
export function preserveCellOnlyWorkbookLayout(source: ArchiveEntries, target: ArchiveEntries): ArchiveEntries {
  const originalSheets = sheets(source), generatedSheets = sheets(target);
  if (originalSheets.length !== generatedSheets.length || originalSheets.some((sheet, index) => sheet.name !== generatedSheets[index].name)) fail("worksheet_structure_changed");
  const output: ArchiveEntries = { ...source };
  const styles = preserveStyleReferences(source, target);
  output["xl/styles.xml"] = styles.bytes;
  for (const [position, originalSheet] of originalSheets.entries()) {
    const sourceSheet = parse(source[originalSheet.path], "source_sheet");
    const generatedSheet = parse(target[generatedSheets[position].path], "generated_sheet");
    const originalData = child(sourceSheet.documentElement, "sheetData"), generatedData = child(generatedSheet.documentElement, "sheetData");
    if (!originalData || !generatedData) fail("missing_sheet_data");
    const data = sourceSheet.importNode(generatedData, true) as Element;
    const rows = new Map(elements(data).map((row) => [row.getAttribute("r"), row]));
    for (const originalRow of elements(originalData)) {
      const id = originalRow.getAttribute("r");
      let row = rows.get(id);
      if (!row) {
        if (!Array.from(originalRow.attributes).some((item) => !["r", "spans"].includes(item.localName))) continue;
        row = sourceSheet.importNode(originalRow.cloneNode(false), true) as Element;
        const next = elements(data).find((item) => Number(item.getAttribute("r")) > Number(id));
        data.insertBefore(row, next ?? null);
      }
      for (const attr of Array.from(row.attributes)) if (attr.localName !== "r") row.removeAttributeNode(attr);
      for (const attr of Array.from(originalRow.attributes)) if (attr.localName !== "spans") row.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
      if (row.hasAttribute("s")) {
        const mapped = styles.originalIds.get(index(row.getAttribute("s")));
        if (mapped === undefined) fail("unmapped_row_style");
        row.setAttribute("s", String(mapped));
      }
    }
    sourceSheet.documentElement.replaceChild(data, originalData);
    const originalDimension = child(sourceSheet.documentElement, "dimension"), generatedDimension = child(generatedSheet.documentElement, "dimension");
    if (originalDimension) sourceSheet.documentElement.removeChild(originalDimension);
    if (generatedDimension) {
      const before = elements(sourceSheet.documentElement).find((element) => element.localName !== "sheetPr");
      sourceSheet.documentElement.insertBefore(sourceSheet.importNode(generatedDimension, true), before ?? null);
    }
    const columns = child(sourceSheet.documentElement, "cols");
    if (columns) for (const column of elements(columns)) if (column.hasAttribute("style")) {
      const mapped = styles.originalIds.get(index(column.getAttribute("style")));
      if (mapped === undefined) fail("unmapped_column_style");
      column.setAttribute("style", String(mapped));
    }
    output[originalSheet.path] = encode(sourceSheet);
  }

  // Generated cell indices address the generated string pool. Add its package
  // references without changing any source relationship IDs or other parts.
  if (target["xl/sharedStrings.xml"]) {
    output["xl/sharedStrings.xml"] = target["xl/sharedStrings.xml"];
    const rels = parse(output["xl/_rels/workbook.xml.rels"], "workbook_relationships");
    const existing = elements(rels.documentElement).find((item) => item.getAttribute("Type") === `${REL}/sharedStrings`);
    if (existing) {
      if (resolveTarget(existing.getAttribute("Target"), "xl/workbook.xml") !== "xl/sharedStrings.xml") fail("nonstandard_shared_strings_path");
    } else {
      const used = new Set(elements(rels.documentElement).map((item) => item.getAttribute("Id")));
      let suffix = 1;
      while (used.has(`rIdCellStrings${suffix}`)) suffix += 1;
      const relationship = rels.createElementNS(PKG_REL, "Relationship");
      relationship.setAttribute("Id", `rIdCellStrings${suffix}`);
      relationship.setAttribute("Type", `${REL}/sharedStrings`);
      relationship.setAttribute("Target", "sharedStrings.xml");
      rels.documentElement.appendChild(relationship);
      output["xl/_rels/workbook.xml.rels"] = encode(rels);
    }
    const types = parse(output["[Content_Types].xml"], "content_types");
    if (!elements(types.documentElement).some((item) => item.getAttribute("PartName") === "/xl/sharedStrings.xml")) {
      const override = types.createElementNS(TYPES, "Override");
      override.setAttribute("PartName", "/xl/sharedStrings.xml");
      override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
      types.documentElement.appendChild(override);
      output["[Content_Types].xml"] = encode(types);
    }
  }
  return output;
}
