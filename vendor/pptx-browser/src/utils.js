/**
 * utils.js — XML parsing helpers and shared constants.
 */

export function parseXml(str) {
  return new DOMParser().parseFromString(str, 'application/xml');
}

/** Get all descendant elements with a given local name (namespace-agnostic). */
export function gtn(node, localName) {
  if (!node) return [];
  const results = [];
  const all = node.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    if (all[i].localName === localName) results.push(all[i]);
  }
  return results;
}

/** Get the first descendant element with the given local name, or null. */
export function g1(node, localName) {
  return gtn(node, localName)[0] ?? null;
}

/** Get an attribute value, or `def` if absent. */
export function attr(el, name, def = null) {
  if (!el) return def;
  const v = el.getAttribute(name);
  return v !== null ? v : def;
}

/** Get an attribute value as an integer, or `def` if absent. */
export function attrInt(el, name, def = 0) {
  const v = attr(el, name);
  return v !== null ? parseInt(v, 10) : def;
}

/** Get an attribute value as a float, or `def` if absent. */
export function attrFloat(el, name, def = 0) {
  const v = attr(el, name);
  return v !== null ? parseFloat(v) : def;
}

export function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

// ── Unit constants ──────────────────────────────────────────────────────────
/** EMU (English Metric Units) per inch */
export const EMU_PER_INCH = 914400;
/** EMU per typographic point */
export const EMU_PER_PT = 12700;
