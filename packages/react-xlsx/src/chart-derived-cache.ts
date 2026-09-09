import { strFromU8, strToU8 } from "fflate";

const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CACHE_NAMES = new Map([
  ["numRef", "numCache"],
  ["strRef", "strCache"],
  ["multiLvlStrRef", "multiLvlStrCache"]
]);

/** Reference caches are derived data; literal chart points remain authoritative. */
export function invalidateDerivedChartCaches(archive: Record<string, Uint8Array>): void {
  for (const path of Object.keys(archive)) {
    if (!/^xl\/charts\/[^/]+\.xml$/i.test(path)) continue;
    const xml = strFromU8(archive[path]);
    if (!/\b(?:numCache|strCache|multiLvlStrCache)\b/.test(xml)) continue;
    if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error("xlsx_chart_cache_invalid_xml");
    const document = new DOMParser().parseFromString(xml, "application/xml");
    if (document.getElementsByTagName("parsererror").length) throw new Error("xlsx_chart_cache_invalid_xml");
    let changed = false;
    for (const [referenceName, cacheName] of CACHE_NAMES) {
      for (const reference of Array.from(document.getElementsByTagNameNS(CHART_NS, referenceName))) {
        const children = Array.from(reference.children);
        if (!children.some((node) => node.namespaceURI === CHART_NS && node.localName === "f" && node.textContent?.trim())) continue;
        for (const cache of children.filter((node) => node.namespaceURI === CHART_NS && node.localName === cacheName)) {
          reference.removeChild(cache);
          changed = true;
        }
      }
    }
    if (changed) archive[path] = strToU8(new XMLSerializer().serializeToString(document));
  }
}
