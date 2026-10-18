/**
 * theme.js — Parse OOXML theme XML and build effective theme colour tables.
 */

import { g1 } from './utils.js';


/**
 * Parse ppt/theme/theme1.xml and return:
 *   { colors, majorFont, minorFont }
 * where `colors` maps scheme names → hex strings.
 */
export function parseTheme(themeDoc) {
  if (!themeDoc) {
    return { colors: {}, majorFont: 'Calibri Light', minorFont: 'Calibri' };
  }

  // ── Colour scheme ──────────────────────────────────────────────────────────
  const clrScheme = g1(themeDoc, 'clrScheme');
  const colors = {};

  if (clrScheme) {
    const slots = ['dk1','lt1','dk2','lt2',
                   'accent1','accent2','accent3','accent4','accent5','accent6',
                   'hlink','folHlink'];
    for (const key of slots) {
      const el = g1(clrScheme, key);
      if (!el) continue;
      const srgb  = g1(el, 'srgbClr');
      const sysClr = g1(el, 'sysClr');
      if (srgb) {
        colors[key] = srgb.getAttribute('val') || '';
      } else if (sysClr) {
        colors[key] = sysClr.getAttribute('lastClr') || '';
      }
    }
  }

  // ── Font scheme ────────────────────────────────────────────────────────────
  const fontScheme = g1(themeDoc, 'fontScheme');
  let majorFont = 'Calibri Light', minorFont = 'Calibri';

  if (fontScheme) {
    const majorFontEl = g1(fontScheme, 'majorFont');
    const minorFontEl = g1(fontScheme, 'minorFont');
    if (majorFontEl) {
      const latin = g1(majorFontEl, 'latin');
      if (latin) majorFont = latin.getAttribute('typeface') || majorFont;
    }
    if (minorFontEl) {
      const latin = g1(minorFontEl, 'latin');
      if (latin) minorFont = latin.getAttribute('typeface') || minorFont;
    }
  }

  return { colors, majorFont, minorFont };
}

/**
 * Parse the <p:clrMap> element from a slide master.
 * Returns a map like { bg1: 'lt1', tx1: 'dk1', … }
 */
export function parseClrMap(masterDoc) {
  if (!masterDoc) return {};
  const clrMap = g1(masterDoc, 'clrMap');
  if (!clrMap) return {};
  const map = {};
  const attrs = ['bg1','tx1','bg2','tx2',
                 'accent1','accent2','accent3','accent4','accent5','accent6',
                 'hlink','folHlink'];
  for (const a of attrs) {
    const v = clrMap.getAttribute(a);
    if (v) map[a] = v;
  }
  return map;
}

/**
 * Build effective theme colours by applying clrMap remapping.
 * e.g. if clrMap says bg1→lt1, then themeColors.bg1 = themeColors.lt1.
 */
export function buildThemeColors(themeData, clrMap) {
  const base = { ...themeData.colors };
  for (const [key, ref] of Object.entries(clrMap)) {
    if (base[ref] !== undefined) base[key] = base[ref];
  }
  return base;
}
