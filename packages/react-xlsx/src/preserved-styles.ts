const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XMLNS = "http://www.w3.org/2000/xmlns/";
const ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];
const MAX_STYLES = 65_490;

function fail(reason: string): never {
  throw new Error(`xlsx_named_style_preservation_failed:${reason}`);
}

function elements(node: Node): Element[] {
  return Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1);
}

function collection(root: Element, name: string): Element | undefined {
  return elements(root).find((element) => element.namespaceURI === NS && element.localName === name);
}

function parse(xml: string): Document {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) fail("unsafe_xml");
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length || document.documentElement?.localName !== "styleSheet" || document.documentElement.namespaceURI !== NS) fail("invalid_styles_xml");
  return document;
}

function signature(node: Node): string {
  if (node.nodeType === 3 || node.nodeType === 4) {
    if (!node.nodeValue?.trim() && node.parentNode && elements(node.parentNode).length) return "";
    return JSON.stringify(node.nodeValue);
  }
  if (node.nodeType !== 1) return "";
  const element = node as Element;
  const attrs = Array.from(element.attributes).filter((attribute) => attribute.namespaceURI !== XMLNS)
    .map((attribute) => [attribute.namespaceURI ?? "", attribute.localName, attribute.value])
    .sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)));
  return JSON.stringify([element.namespaceURI, element.localName, attrs, Array.from(element.childNodes).map(signature).filter(Boolean)]);
}

function index(value: string | null, label: string): number {
  const number = value === null || value === "" ? 0 : Number(value);
  if (!Number.isSafeInteger(number) || number < 0 || number > 65_535) fail(label);
  return number;
}

function ensure(document: Document, name: string): Element {
  const existing = collection(document.documentElement, name);
  if (existing) return existing;
  const created = document.createElementNS(NS, name);
  const next = elements(document.documentElement).find((element) => ORDER.indexOf(element.localName) > ORDER.indexOf(name));
  document.documentElement.insertBefore(created, next ?? null);
  return created;
}

/**
 * Restore named style definitions without replacing or renumbering generated
 * cell styles. Cell XFs keep their target xfId indexes, whose named definitions
 * are restored by name. Additional named XFs are deduplicated or appended.
 */
export function preserveNamedWorkbookStyles(sourceStylesXml: string, targetStylesXml: string): string {
  const source = parse(sourceStylesXml);
  const target = parse(targetStylesXml);
  const sourceNames = collection(source.documentElement, "cellStyles");
  const sourceXfs = collection(source.documentElement, "cellStyleXfs");
  if (!sourceNames && !sourceXfs) return targetStylesXml;
  if (!sourceXfs) fail("missing_source_named_xfs");
  const componentMaps = new Map<string, Map<number, number>>();

  for (const name of ["fonts", "fills", "borders", "numFmts"]) {
    const original = collection(source.documentElement, name);
    if (!original) {
      if (name !== "numFmts") fail(`missing_${name}`);
      continue;
    }
    const destination = ensure(target, name);
    const existing = elements(destination);
    const bySignature = new Map(existing.map((element, offset) => [signature(element), offset]));
    const mappings = new Map<number, number>();
    componentMaps.set(name, mappings);
    let nextFormatId = Math.max(163, ...existing.map((element) => index(element.getAttribute("numFmtId"), "numFmtId"))) + 1;
    for (const [offset, element] of elements(original).entries()) {
      let copy: Element;
      if (name === "numFmts") {
        const sourceId = index(element.getAttribute("numFmtId"), "numFmtId");
        if (mappings.has(sourceId)) fail("duplicate_number_format");
        // SheetJS writes localized format 56 explicitly even when unused.
        // Keep declarations in the built-in range at their original ID so
        // generated cells using that built-in ID retain its exact definition.
        if (sourceId < 164) {
          const reserved = existing.find((entry) => index(entry.getAttribute("numFmtId"), "numFmtId") === sourceId);
          if (reserved && reserved.getAttribute("formatCode") !== element.getAttribute("formatCode")) fail("conflicting_builtin_number_format");
          mappings.set(sourceId, sourceId);
          if (reserved) continue;
          copy = target.importNode(element, true) as Element;
          if (existing.length >= MAX_STYLES) fail("style_limit");
          destination.appendChild(copy);
          existing.push(copy);
          continue;
        }
        const matching = existing.find((entry) => entry.getAttribute("formatCode") === element.getAttribute("formatCode"));
        if (matching) { mappings.set(sourceId, index(matching.getAttribute("numFmtId"), "numFmtId")); continue; }
        if (nextFormatId > 65_535) fail("custom_number_format_limit");
        copy = target.importNode(element, true) as Element;
        copy.setAttribute("numFmtId", String(nextFormatId));
        mappings.set(sourceId, nextFormatId++);
      } else {
        const key = signature(element);
        const matching = bySignature.get(key);
        if (matching !== undefined) { mappings.set(offset, matching); continue; }
        mappings.set(offset, existing.length);
        bySignature.set(key, existing.length);
        copy = target.importNode(element, true) as Element;
      }
      if (existing.length >= MAX_STYLES) fail("style_limit");
      destination.appendChild(copy);
      existing.push(copy);
    }
    destination.setAttribute("count", String(existing.length));
  }

  const destinationXfs = ensure(target, "cellStyleXfs");
  const targetXfs = elements(destinationXfs);
  const originalXfs = elements(sourceXfs);
  const xfMappings = new Map<number, number>();
  const byXfSignature = new Map(targetXfs.map((xf, offset) => [signature(xf), offset]));
  const preferredSlots = new Map<number, Set<number>>();
  const claimedSlots = new Map<number, number>();
  const targetNames = collection(target.documentElement, "cellStyles");
  if (sourceNames && targetNames) for (const sourceName of elements(sourceNames)) {
    const targetName = elements(targetNames).find((entry) => entry.getAttribute("name") === sourceName.getAttribute("name"));
    if (!targetName) continue;
    const originalId = index(sourceName.getAttribute("xfId"), "xfId");
    const targetId = index(targetName.getAttribute("xfId"), "xfId");
    if (!originalXfs[originalId] || !targetXfs[targetId]) fail("unresolved_named_style");
    if (claimedSlots.has(targetId) && claimedSlots.get(targetId) !== originalId) fail("ambiguous_generated_named_style");
    claimedSlots.set(targetId, originalId);
    const slots = preferredSlots.get(originalId) ?? new Set<number>();
    slots.add(targetId);
    preferredSlots.set(originalId, slots);
  }
  // Restore referenced names first so deduplication of unused original XFs
  // cannot bind to a generated slot that is about to receive its source name.
  const orderedXfs = [...originalXfs.entries()].sort(([left], [right]) => Number(preferredSlots.has(right)) - Number(preferredSlots.has(left)));
  for (const [offset, xf] of orderedXfs) {
    // A named XF is already the root definition. An unexpected parent or
    // extension Id cannot be interpreted safely as a component-table index.
    if (xf.hasAttribute("xfId") || Array.from(xf.attributes).some((attribute) => /Id$/.test(attribute.localName) && !["numFmtId", "fontId", "fillId", "borderId"].includes(attribute.localName))) fail("unsupported_named_xf_reference");
    const copy = target.importNode(xf, true) as Element;
    for (const [attribute, name] of [["fontId", "fonts"], ["fillId", "fills"], ["borderId", "borders"], ["numFmtId", "numFmts"]]) {
      const sourceId = index(xf.getAttribute(attribute), attribute);
      if (attribute === "numFmtId" && sourceId < 164) continue;
      const mapped = componentMaps.get(name)?.get(sourceId);
      if (mapped === undefined) fail(`unresolved_${attribute}`);
      copy.setAttribute(attribute, String(mapped));
    }
    const key = signature(copy);
    const preferred = preferredSlots.get(offset);
    if (preferred?.size) {
      for (const slot of preferred) {
        const replacement = target.importNode(copy, true) as Element;
        destinationXfs.replaceChild(replacement, targetXfs[slot]);
        targetXfs[slot] = replacement;
      }
      xfMappings.set(offset, preferred.values().next().value!);
      byXfSignature.clear();
      targetXfs.forEach((entry, position) => byXfSignature.set(signature(entry), position));
      continue;
    }
    const matching = byXfSignature.get(key);
    if (matching !== undefined) { xfMappings.set(offset, matching); continue; }
    if (targetXfs.length >= MAX_STYLES) fail("named_style_limit");
    xfMappings.set(offset, targetXfs.length);
    byXfSignature.set(key, targetXfs.length);
    destinationXfs.appendChild(copy);
    targetXfs.push(copy);
  }
  destinationXfs.setAttribute("count", String(targetXfs.length));

  // Keep generated names that have no source counterpart, while restoring every
  // source name and all its attributes (built-in identity, hidden state, etc.).
  if (sourceNames) {
    const destinationNames = ensure(target, "cellStyles");
    const names = new Set<string>();
    for (const style of elements(sourceNames)) {
      const name = style.getAttribute("name");
      if (!name || names.has(name)) fail("duplicate_or_missing_style_name");
      names.add(name);
      const mapped = xfMappings.get(index(style.getAttribute("xfId"), "xfId"));
      if (mapped === undefined) fail("unresolved_named_style");
      const copy = target.importNode(style, true) as Element;
      copy.setAttribute("xfId", String(mapped));
      const existing = elements(destinationNames).find((entry) => entry.getAttribute("name") === name);
      if (existing) destinationNames.replaceChild(copy, existing);
      else destinationNames.appendChild(copy);
    }
    destinationNames.setAttribute("count", String(elements(destinationNames).length));
  }
  const cellXfs = collection(target.documentElement, "cellXfs");
  if (cellXfs) for (const cell of elements(cellXfs)) if (index(cell.getAttribute("xfId"), "xfId") >= targetXfs.length) fail("unresolved_generated_cell_xf");
  return new XMLSerializer().serializeToString(target);
}
