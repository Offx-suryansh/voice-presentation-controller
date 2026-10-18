/**
 * fonts.js — Font resolution, loading, and custom registration.
 *
 * Priority order for font resolution (highest first):
 *  1. User-registered custom fonts (registerFont)
 *  2. Fonts embedded in the PPTX file (extractEmbeddedFonts)
 *  3. Already-loaded system/document fonts (checked via FontFaceSet)
 *  4. MS Office → Google Fonts mapping (loaded on demand)
 *  5. Direct Google Fonts attempt for unknown fonts
 *  6. Generic family fallback (sans-serif / serif / monospace)
 */

import { g1 } from './utils.js';

// ── MS Office → Google Fonts mapping ────────────────────────────────────────
// key   = lowercase MS font name (exact match)
// value = { google: string|null, weights: number[], generic: string }
//   google:  null  → web-safe, use as-is (no loading needed)
//            string → Google Fonts family name to load instead
//   generic: fallback CSS generic family

const MS_FONT_MAP = {
  // ── Calibri family (metric-compatible substitutes exist) ──────────────────
  'calibri':                { google: 'Carlito',            weights: [400,700],      generic: 'sans-serif' },
  'calibri light':          { google: 'Carlito',            weights: [300],          generic: 'sans-serif' },
  'calibri (body)':         { google: 'Carlito',            weights: [400,700],      generic: 'sans-serif' },

  // ── Cambria (metric-compatible) ───────────────────────────────────────────
  'cambria':                { google: 'Caladea',            weights: [400,700],      generic: 'serif' },
  'cambria math':           { google: 'Caladea',            weights: [400],          generic: 'serif' },

  // ── Aptos — Microsoft 365 default since 2023 ─────────────────────────────
  'aptos':                  { google: 'Inter',              weights: [300,400,600,700], generic: 'sans-serif' },
  'aptos display':          { google: 'Inter',              weights: [700,800],      generic: 'sans-serif' },
  'aptos narrow':           { google: 'Inter',              weights: [400,700],      generic: 'sans-serif' },
  'aptos serif':            { google: 'Lora',               weights: [400,700],      generic: 'serif' },
  'aptos mono':             { google: 'Roboto Mono',        weights: [400,700],      generic: 'monospace' },

  // ── Web-safe fonts — available in all browsers, no loading needed ─────────
  'arial':                  { google: null, weights: [], generic: 'sans-serif' },
  'arial black':            { google: null, weights: [], generic: 'sans-serif' },
  'times new roman':        { google: null, weights: [], generic: 'serif' },
  'times':                  { google: null, weights: [], generic: 'serif' },
  'helvetica':              { google: null, weights: [], generic: 'sans-serif' },
  'verdana':                { google: null, weights: [], generic: 'sans-serif' },
  'tahoma':                 { google: null, weights: [], generic: 'sans-serif' },
  'trebuchet ms':           { google: null, weights: [], generic: 'sans-serif' },
  'georgia':                { google: null, weights: [], generic: 'serif' },
  'courier new':            { google: null, weights: [], generic: 'monospace' },
  'courier':                { google: null, weights: [], generic: 'monospace' },
  'impact':                 { google: null, weights: [], generic: 'sans-serif' },
  'comic sans ms':          { google: null, weights: [], generic: 'cursive' },
  'palatino':               { google: null, weights: [], generic: 'serif' },
  'lucida console':         { google: null, weights: [], generic: 'monospace' },
  'lucida sans unicode':    { google: null, weights: [], generic: 'sans-serif' },

  // ── Common Office fonts ───────────────────────────────────────────────────
  'arial narrow':           { google: 'Arimo',              weights: [400,700],      generic: 'sans-serif' },
  'candara':                { google: 'Nunito',             weights: [300,400,700],  generic: 'sans-serif' },
  'consolas':               { google: 'Roboto Mono',        weights: [400,700],      generic: 'monospace' },
  'constantia':             { google: 'Libre Baskerville',  weights: [400,700],      generic: 'serif' },
  'corbel':                 { google: 'Lato',               weights: [300,400,700],  generic: 'sans-serif' },
  'franklin gothic medium': { google: 'Libre Franklin',     weights: [500],          generic: 'sans-serif' },
  'franklin gothic book':   { google: 'Libre Franklin',     weights: [400],          generic: 'sans-serif' },
  'franklin gothic heavy':  { google: 'Libre Franklin',     weights: [800],          generic: 'sans-serif' },
  'gill sans mt':           { google: 'Quattrocento Sans',  weights: [400,700],      generic: 'sans-serif' },
  'gill sans':              { google: 'Quattrocento Sans',  weights: [400,700],      generic: 'sans-serif' },
  'century gothic':         { google: 'Josefin Sans',       weights: [300,400,700],  generic: 'sans-serif' },
  'century schoolbook':     { google: 'EB Garamond',        weights: [400,700],      generic: 'serif' },
  'garamond':               { google: 'EB Garamond',        weights: [400,700],      generic: 'serif' },
  'palatino linotype':      { google: 'EB Garamond',        weights: [400,700],      generic: 'serif' },
  'book antiqua':           { google: 'EB Garamond',        weights: [400,700],      generic: 'serif' },
  'rockwell':               { google: 'Roboto Slab',        weights: [400,700],      generic: 'serif' },
  'rockwell extra bold':    { google: 'Roboto Slab',        weights: [800],          generic: 'serif' },
  'segoe ui':               { google: 'Inter',              weights: [300,400,600,700], generic: 'sans-serif' },
  'segoe ui light':         { google: 'Inter',              weights: [300],          generic: 'sans-serif' },
  'segoe ui semibold':      { google: 'Inter',              weights: [600],          generic: 'sans-serif' },
  'segoe ui semilight':     { google: 'Inter',              weights: [350],          generic: 'sans-serif' },
  'helvetica neue':         { google: 'Nunito Sans',        weights: [300,400,700],  generic: 'sans-serif' },
  'myriad pro':             { google: 'Source Sans 3',      weights: [300,400,600,700], generic: 'sans-serif' },
  'futura':                 { google: 'Josefin Sans',       weights: [300,400,700],  generic: 'sans-serif' },
  'tw cen mt':              { google: 'Pathway Gothic One', weights: [400],          generic: 'sans-serif' },
  'bookman old style':      { google: 'Libre Baskerville',  weights: [400,700],      generic: 'serif' },
  'frutiger':               { google: 'Raleway',            weights: [300,400,700],  generic: 'sans-serif' },
  'optima':                 { google: 'Questrial',          weights: [400],          generic: 'sans-serif' },
  'univers':                { google: 'Nunito Sans',        weights: [300,400,700],  generic: 'sans-serif' },

  // ── Google Fonts already at their canonical names ─────────────────────────
  'open sans':              { google: 'Open Sans',          weights: [300,400,600,700], generic: 'sans-serif' },
  'lato':                   { google: 'Lato',               weights: [300,400,700],  generic: 'sans-serif' },
  'montserrat':             { google: 'Montserrat',         weights: [300,400,600,700], generic: 'sans-serif' },
  'raleway':                { google: 'Raleway',            weights: [300,400,700],  generic: 'sans-serif' },
  'roboto':                 { google: 'Roboto',             weights: [300,400,700],  generic: 'sans-serif' },
  'roboto mono':            { google: 'Roboto Mono',        weights: [400,700],      generic: 'monospace' },
  'roboto slab':            { google: 'Roboto Slab',        weights: [300,400,700],  generic: 'serif' },
  'oswald':                 { google: 'Oswald',             weights: [400,700],      generic: 'sans-serif' },
  'playfair display':       { google: 'Playfair Display',   weights: [400,700],      generic: 'serif' },
  'merriweather':           { google: 'Merriweather',       weights: [300,400,700],  generic: 'serif' },
  'nunito':                 { google: 'Nunito',             weights: [300,400,700],  generic: 'sans-serif' },
  'nunito sans':            { google: 'Nunito Sans',        weights: [300,400,700],  generic: 'sans-serif' },
  'poppins':                { google: 'Poppins',            weights: [300,400,600,700], generic: 'sans-serif' },
  'inter':                  { google: 'Inter',              weights: [300,400,600,700], generic: 'sans-serif' },
  'work sans':              { google: 'Work Sans',          weights: [300,400,600,700], generic: 'sans-serif' },
  'dm sans':                { google: 'DM Sans',            weights: [300,400,500,700], generic: 'sans-serif' },
  'dm serif display':       { google: 'DM Serif Display',   weights: [400],          generic: 'serif' },
  'ubuntu':                 { google: 'Ubuntu',             weights: [300,400,700],  generic: 'sans-serif' },
  'ubuntu mono':            { google: 'Ubuntu Mono',        weights: [400,700],      generic: 'monospace' },
  'source sans pro':        { google: 'Source Sans 3',      weights: [300,400,600,700], generic: 'sans-serif' },
  'source serif pro':       { google: 'Source Serif 4',     weights: [300,400,700],  generic: 'serif' },
  'source code pro':        { google: 'Source Code Pro',    weights: [400,700],      generic: 'monospace' },
  'exo 2':                  { google: 'Exo 2',              weights: [300,400,700],  generic: 'sans-serif' },
  'titillium web':          { google: 'Titillium Web',      weights: [300,400,600,700], generic: 'sans-serif' },
  'fira sans':              { google: 'Fira Sans',          weights: [300,400,600,700], generic: 'sans-serif' },
  'fira mono':              { google: 'Fira Mono',          weights: [400,700],      generic: 'monospace' },
  'josefin sans':           { google: 'Josefin Sans',       weights: [300,400,700],  generic: 'sans-serif' },
  'josefin slab':           { google: 'Josefin Slab',       weights: [300,400,700],  generic: 'serif' },
  'barlow':                 { google: 'Barlow',             weights: [300,400,600,700], generic: 'sans-serif' },
  'barlow condensed':       { google: 'Barlow Condensed',   weights: [300,400,600,700], generic: 'sans-serif' },
  'cabin':                  { google: 'Cabin',              weights: [400,700],      generic: 'sans-serif' },
  'crimson text':           { google: 'Crimson Text',       weights: [400,700],      generic: 'serif' },
  'libre baskerville':      { google: 'Libre Baskerville',  weights: [400,700],      generic: 'serif' },
  'libre franklin':         { google: 'Libre Franklin',     weights: [400,700],      generic: 'sans-serif' },
  'eb garamond':            { google: 'EB Garamond',        weights: [400,700],      generic: 'serif' },
  'spectral':               { google: 'Spectral',           weights: [300,400,700],  generic: 'serif' },
  'arvo':                   { google: 'Arvo',               weights: [400,700],      generic: 'serif' },
  'pt sans':                { google: 'PT Sans',            weights: [400,700],      generic: 'sans-serif' },
  'pt serif':               { google: 'PT Serif',           weights: [400,700],      generic: 'serif' },
  'pt mono':                { google: 'PT Mono',            weights: [400],          generic: 'monospace' },
  'karla':                  { google: 'Karla',              weights: [300,400,700],  generic: 'sans-serif' },
  'mukta':                  { google: 'Mukta',              weights: [300,400,700],  generic: 'sans-serif' },
  'hind':                   { google: 'Hind',               weights: [300,400,700],  generic: 'sans-serif' },
  'noto sans':              { google: 'Noto Sans',          weights: [300,400,700],  generic: 'sans-serif' },
  'noto serif':             { google: 'Noto Serif',         weights: [300,400,700],  generic: 'serif' },
};

// ── Internal state ──────────────────────────────────────────────────────────

/** Set of font family names already loaded (prevents duplicate requests). */
const _loadedFonts = new Set();

/** User-registered custom fonts: family name → array of FontFace objects. */
const _customFonts = new Map();

/** Fonts confirmed available on this system via font-check. */
const _systemFonts = new Set();

// ── Canvas font-availability probe ──────────────────────────────────────────

/** Shared measurement canvas — created once and reused. */
let _probeCanvas = null;
let _probeCtx = null;

function getProbeCtx() {
  if (!_probeCtx) {
    _probeCanvas = typeof OffscreenCanvas !== 'undefined'
      ? new OffscreenCanvas(300, 10)
      : Object.assign(document.createElement('canvas'), { width: 300, height: 10 });
    _probeCtx = _probeCanvas.getContext('2d');
  }
  return _probeCtx;
}

const PROBE_TEXT  = 'mmmmmmmmmmlllllllllliiiiiiiiiixxxxxxxxxx';
const PROBE_SIZES = [20, 40]; // measure at 2 sizes to avoid coincidental matches

/**
 * Check whether a font family is available in this browser/system.
 * Uses a double-size measurement trick: if the font metrics differ from both
 * monospace and serif, the font is available.
 *
 * Returns true if the font appears to be present.
 */
export function isFontAvailable(family) {
  if (_systemFonts.has(family)) return true;

  // If FontFaceSet has it, it's definitely loaded
  if (typeof document !== 'undefined' && document.fonts) {
    if (document.fonts.check(`16px "${family}"`)) {
      _systemFonts.add(family);
      return true;
    }
  }

  // Fallback: canvas measurement trick (compare against two baseline fonts)
  try {
    const ctx = getProbeCtx();
    const baselines = ['monospace', 'serif'];

    for (const size of PROBE_SIZES) {
      for (const baseline of baselines) {
        ctx.font = `${size}px ${baseline}`;
        const baseW = ctx.measureText(PROBE_TEXT).width;
        ctx.font = `${size}px "${family}", ${baseline}`;
        const testW = ctx.measureText(PROBE_TEXT).width;
        if (Math.abs(testW - baseW) > 0.5) {
          _systemFonts.add(family);
          return true;
        }
      }
    }
  } catch (_) {}

  return false;
}

// ── Custom font registration ─────────────────────────────────────────────────

/**
 * Register a custom font so the renderer uses it for matching typeface names
 * in the PPTX. You can register the same family multiple times with different
 * descriptors to provide bold/italic variants.
 *
 * @param {string} family
 *   The exact font family name as used in the PPTX (e.g. "Brand Sans")
 *   OR the MS Office name it should replace (e.g. "Calibri")
 * @param {string | URL | File | ArrayBuffer | Uint8Array} source
 *   Where to load the font from:
 *     - string / URL  → a URL (https:// or data:)
 *     - File          → a File object from <input type="file">
 *     - ArrayBuffer / Uint8Array → raw font bytes (ttf / woff / woff2 / otf)
 * @param {FontFaceDescriptors} [descriptors]
 *   Optional FontFace descriptors: { weight, style, unicodeRange, … }
 *   Defaults to { weight: 'normal', style: 'normal' }
 * @returns {Promise<FontFace>} the registered FontFace
 *
 * @example
 * // From a URL
 * await renderer.registerFont('Brand Sans', '/fonts/brand-sans.woff2');
 *
 * // Bold variant
 * await renderer.registerFont('Brand Sans', '/fonts/brand-sans-bold.woff2', { weight: '700' });
 *
 * // From a File input
 * const [file] = e.target.files;
 * await renderer.registerFont('Brand Sans', file);
 *
 * // Override Calibri globally
 * await renderer.registerFont('Calibri', '/fonts/my-calibri.woff2');
 */
export async function registerFont(family, source, descriptors = {}) {
  let fontSource;

  if (typeof source === 'string' || source instanceof URL) {
    fontSource = source.toString();
  } else if (source instanceof File || source instanceof Blob) {
    const buf = await source.arrayBuffer();
    fontSource = buf;
  } else if (source instanceof Uint8Array) {
    fontSource = source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
  } else if (source instanceof ArrayBuffer) {
    fontSource = source;
  } else {
    throw new TypeError(`registerFont: unsupported source type. Expected string, URL, File, Blob, ArrayBuffer, or Uint8Array.`);
  }

  const desc = { weight: 'normal', style: 'normal', ...descriptors };
  const face = new FontFace(family, typeof fontSource === 'string' ? `url(${fontSource})` : fontSource, desc);

  await face.load();
  document.fonts.add(face);

  // Track in our registry
  const variants = _customFonts.get(family) ?? [];
  variants.push(face);
  _customFonts.set(family, variants);

  // Mark as loaded so we skip the Google Fonts request for this name
  _loadedFonts.add(family);
  _systemFonts.add(family);

  // Also register under the lowercase key so resolveFontFamily finds it
  const lower = family.toLowerCase().trim();
  if (!MS_FONT_MAP[lower]) {
    MS_FONT_MAP[lower] = { google: null, weights: [], generic: 'sans-serif', _custom: true };
  } else {
    MS_FONT_MAP[lower]._custom = true;
  }

  console.log(`[pptx-canvas-renderer] Custom font registered: "${family}" (${desc.weight} ${desc.style})`);
  return face;
}

/**
 * Register multiple fonts at once from an object map.
 * @param {Record<string, string | { url: string, weight?: string, style?: string }[]>} fontMap
 *
 * @example
 * await renderer.registerFonts({
 *   'Brand Sans': '/fonts/brand-sans.woff2',
 *   'Brand Serif': [
 *     { url: '/fonts/brand-serif.woff2', weight: '400' },
 *     { url: '/fonts/brand-serif-bold.woff2', weight: '700' },
 *   ]
 * });
 */
export async function registerFonts(fontMap) {
  const promises = [];
  for (const [family, spec] of Object.entries(fontMap)) {
    if (typeof spec === 'string') {
      promises.push(registerFont(family, spec));
    } else if (Array.isArray(spec)) {
      for (const variant of spec) {
        const { url, ...desc } = variant;
        promises.push(registerFont(family, url, desc));
      }
    }
  }
  await Promise.all(promises);
}

// ── Embedded PPTX fonts ──────────────────────────────────────────────────────

/**
 * Detect fonts embedded in the PPTX file (stored in ppt/fonts/).
 *
 * PPTX embeds fonts as .fntdata files — a proprietary Microsoft format that
 * cannot be used directly as a web font. This function detects which fonts
 * are embedded and returns diagnostic info so the renderer can decide what to do.
 *
 * @param {Document}  presDoc   — parsed presentation.xml
 * @param {object}    presRels  — relationship map from getRels()
 * @returns {EmbeddedFontInfo[]}
 *
 * @typedef {object} EmbeddedFontInfo
 * @property {string}  family    — font family name
 * @property {boolean} hasRegular
 * @property {boolean} hasBold
 * @property {boolean} hasItalic
 * @property {boolean} hasBoldItalic
 * @property {string[]} paths    — ZIP paths to .fntdata files (not directly usable as web fonts)
 */
export function detectEmbeddedFonts(presDoc, presRels) {
  if (!presDoc) return [];
  const embeddedFontLst = g1(presDoc, 'embeddedFontLst');
  if (!embeddedFontLst) return [];

  const result = [];

  for (const embeddedFont of embeddedFontLst.children) {
    if (embeddedFont.localName !== 'embeddedFont') continue;

    const fontEl = g1(embeddedFont, 'font');
    const family = fontEl ? fontEl.getAttribute('typeface') : null;
    if (!family) continue;

    const variants = { regular: false, bold: false, italic: false, boldItalic: false };
    const paths = [];

    for (const variant of ['regular', 'bold', 'italic', 'boldItalic']) {
      const el = g1(embeddedFont, variant);
      if (!el) continue;
      const rId = el.getAttribute('r:id') || el.getAttribute('id');
      const rel = presRels?.[rId];
      if (rel) {
        paths.push(rel.fullPath);
        variants[variant] = true;
      }
    }

    result.push({
      family,
      hasRegular:    variants.regular,
      hasBold:       variants.bold,
      hasItalic:     variants.italic,
      hasBoldItalic: variants.boldItalic,
      paths,
      note: '.fntdata files are a proprietary Microsoft format and cannot be used as web fonts directly. Use registerFont() with a compatible woff2/ttf version instead.',
    });
  }

  return result;
}

// ── Font resolution ──────────────────────────────────────────────────────────

/**
 * Resolve a font name to its best available web equivalent.
 *
 * Priority:
 *  1. Custom registered font (user-provided or embedded)
 *  2. Already available in document.fonts / system
 *  3. MS→Google mapping
 *  4. Same name (hope it's available on the system)
 *
 * Returns the final CSS font family name to use.
 */
export function resolveFontFamily(name) {
  if (!name) return 'sans-serif';

  // Theme font tokens — resolved at call site via themeData
  if (name === '+mj-lt' || name === '+mj') return 'serif';
  if (name === '+mn-lt' || name === '+mn') return 'sans-serif';

  // Custom registered font takes top priority
  if (_customFonts.has(name)) return name;

  const lower = name.toLowerCase().trim();
  const mapped = MS_FONT_MAP[lower];

  if (mapped) {
    // Custom override for a known MS font name
    if (mapped._custom) return name;
    // Web-safe — use as-is
    if (mapped.google === null) return name;
    // Check if the substitute is available (may already be on the system)
    if (mapped.google && isFontAvailable(mapped.google)) return mapped.google;
    // Return the mapped name (it will be loaded by loadGoogleFontsFor)
    return mapped.google || name;
  }

  // Unknown font — return as-is; may be available on the system
  return name;
}

/**
 * Get the generic family fallback for a font name.
 * Used to build robust CSS font stacks.
 */
export function getGenericFamily(name) {
  if (!name) return 'sans-serif';
  const lower = name.toLowerCase().trim();
  return MS_FONT_MAP[lower]?.generic ?? 'sans-serif';
}

// ── Google Fonts loader ──────────────────────────────────────────────────────

/**
 * Load Google Fonts substitutes for any MS Office fonts used in the slide.
 * Skips fonts that are:
 *  - Already loaded (tracked in _loadedFonts)
 *  - Web-safe (google: null)
 *  - User-registered custom fonts
 *  - Already available on the system
 *
 * @param {Set<string>} fontNames  — raw font names from the PPTX
 * @param {object}      themeData  — { majorFont, minorFont }
 */
export async function loadGoogleFontsFor(fontNames, themeData) {
  // Collect candidates (font names + theme fonts)
  const candidates = [...fontNames];
  if (themeData?.majorFont) candidates.push(themeData.majorFont);
  if (themeData?.minorFont) candidates.push(themeData.minorFont);

  // Map: googleFamilyName → Set<weight>
  const toLoad = new Map();

  for (const name of candidates) {
    if (!name || name.startsWith('+')) continue;

    // Skip theme tokens
    if (name === '+mj-lt' || name === '+mn-lt' || name === '+mj' || name === '+mn') continue;

    // Skip custom-registered fonts
    if (_customFonts.has(name)) continue;

    const lower = name.toLowerCase().trim();
    const mapped = MS_FONT_MAP[lower];

    let googleName = null;
    let weights = [400, 700];

    if (mapped) {
      if (mapped._custom) continue;          // user-registered override
      if (mapped.google === null) continue;  // web-safe
      googleName = mapped.google;
      weights = mapped.weights.length ? mapped.weights : [400, 700];
    } else {
      // Unknown font — try to load it directly from Google Fonts
      // (works for fonts like "Pacifico", "Dancing Script", etc.)
      googleName = name;
    }

    if (!googleName) continue;

    // Skip if already loaded or confirmed system-available
    if (_loadedFonts.has(googleName)) continue;
    if (isFontAvailable(googleName)) {
      _loadedFonts.add(googleName);
      continue;
    }

    const ws = toLoad.get(googleName) ?? new Set();
    weights.forEach(w => ws.add(w));
    toLoad.set(googleName, ws);
  }

  if (toLoad.size === 0) return;

  // Mark as pending immediately (prevents duplicate concurrent requests)
  for (const name of toLoad.keys()) _loadedFonts.add(name);

  // Build a single batched Google Fonts CSS2 request
  // Format: family=Name:ital,wght@0,400;0,700;1,400;1,700
  const params = [...toLoad.entries()].map(([family, weightSet]) => {
    const wArr = [...weightSet].sort((a, b) => a - b);
    const specs = [
      ...wArr.map(w => `0,${w}`),
      ...wArr.map(w => `1,${w}`),
    ].join(';');
    return `family=${encodeURIComponent(`${family}:ital,wght@${specs}`)}`;
  });

  const url = `https://fonts.googleapis.com/css2?${params.join('&')}&display=swap`;

  try {
    const link = document.createElement('link');
    link.rel  = 'stylesheet';
    link.href = url;
    document.head.appendChild(link);

    // Wait for stylesheet to load
    await new Promise(resolve => {
      link.onload  = resolve;
      link.onerror = () => {
        console.warn('[pptx-canvas-renderer] Google Fonts failed to load:', url);
        resolve();
      };
      setTimeout(resolve, 5000); // hard timeout
    });

    // Let the browser parse and prepare all font faces
    if (document.fonts?.ready) {
      await Promise.race([
        document.fonts.ready,
        new Promise(r => setTimeout(r, 2000)),
      ]);
    }
  } catch (err) {
    console.warn('[pptx-canvas-renderer] Font loading error:', err);
  }
}

// ── Font collection ──────────────────────────────────────────────────────────

/**
 * Scan a set of parsed XML documents and collect all distinct font names
 * referenced in run properties (<a:latin>, <a:ea>, <a:cs> elements).
 *
 * @param {Document[]} xmlDocs
 * @returns {Set<string>}
 */
export function collectUsedFonts(xmlDocs) {
  const names = new Set();
  for (const doc of xmlDocs) {
    if (!doc) continue;
    const els = doc.getElementsByTagName('*');
    for (let i = 0; i < els.length; i++) {
      const el = els[i];
      const ln = el.localName;
      if (ln === 'latin' || ln === 'ea' || ln === 'cs') {
        const tf = el.getAttribute('typeface');
        // Skip theme font tokens ('+mj-lt', '+mn-lt', etc.)
        if (tf && !tf.startsWith('+')) names.add(tf);
      }
    }
  }
  return names;
}

// ── Font building ────────────────────────────────────────────────────────────

/**
 * Build a canvas font string using the full OOXML inheritance chain.
 *
 * Inheritance order (lowest → highest priority):
 *   lstStyle defRPr → paragraph pPr defRPr → run rPr
 *
 * @param {Element|null} rPr        — run properties (<a:rPr>)
 * @param {Element|null} paraDefRPr — paragraph default run properties (<a:defRPr> in <a:pPr>)
 * @param {number}       scaledPxPerEmu — combined scale factor (px per EMU)
 * @param {object}       themeData  — { majorFont, minorFont }
 * @param {number}       [defSz=1800]   — default font size in 100ths of a point
 * @param {Element|null} [lstDefRPr]    — list style default run properties
 * @returns {{ fontStr, sz, szPx, bold, italic, family, generic }}
 */
export function buildFontInherited(rPr, paraDefRPr, scaledPxPerEmu, themeData, defSz = 1800, lstDefRPr = null) {
  // ── Size (100ths of a point) ───────────────────────────────────────────────
  let sz = defSz;
  if (lstDefRPr) { const v = lstDefRPr.getAttribute('sz'); if (v) sz = parseInt(v, 10); }
  if (paraDefRPr){ const v = paraDefRPr.getAttribute('sz'); if (v) sz = parseInt(v, 10); }
  if (rPr)       { const v = rPr.getAttribute('sz');        if (v) sz = parseInt(v, 10); }

  // ── Bold / italic ─────────────────────────────────────────────────────────
  // Explicit '0' beats inherited '1'
  let bold = false, italic = false;
  if (lstDefRPr) {
    if (lstDefRPr.getAttribute('b') === '1') bold   = true;
    if (lstDefRPr.getAttribute('i') === '1') italic = true;
  }
  if (paraDefRPr) {
    const b = paraDefRPr.getAttribute('b');
    const i = paraDefRPr.getAttribute('i');
    if (b === '1') bold   = true; else if (b === '0') bold   = false;
    if (i === '1') italic = true; else if (i === '0') italic = false;
  }
  if (rPr) {
    const b = rPr.getAttribute('b');
    const i = rPr.getAttribute('i');
    if (b === '1') bold   = true; else if (b === '0') bold   = false;
    if (i === '1') italic = true; else if (i === '0') italic = false;
  }

  // ── Font family ───────────────────────────────────────────────────────────
  let rawFamily = themeData?.minorFont ?? 'Calibri';

  function applyFamilyFromEl(el) {
    if (!el) return;
    // Try <a:latin> child first, then typeface attr on the element itself
    const latin = g1(el, 'latin');
    const tf    = latin ? latin.getAttribute('typeface') : el.getAttribute('typeface');
    if (!tf) return;
    if (tf === '+mj-lt' || tf === '+mj') { rawFamily = themeData?.majorFont ?? rawFamily; return; }
    if (tf === '+mn-lt' || tf === '+mn') { rawFamily = themeData?.minorFont ?? rawFamily; return; }
    rawFamily = tf;
  }

  applyFamilyFromEl(lstDefRPr);
  applyFamilyFromEl(paraDefRPr);
  applyFamilyFromEl(rPr);

  const family  = resolveFontFamily(rawFamily);
  const generic = getGenericFamily(rawFamily);

  // ── Build canvas font string ──────────────────────────────────────────────
  // sz (100ths of pt) → px: sz / 100 pt × 12700 EMU/pt × scaledPxPerEmu px/EMU
  //                    = sz × 127 × scaledPxPerEmu
  const szPx    = sz * 127 * scaledPxPerEmu;
  const weight  = bold ? 'bold' : 'normal';
  const style   = italic ? 'italic ' : '';
  // Build a font stack: preferred font → generic family
  const fontStr = `${style}${weight} ${szPx}px "${family}", ${generic}`;

  return { fontStr, sz, szPx, bold, italic, family, generic, rawFamily };
}

/**
 * Simpler font builder for cases without inheritance.
 */
export function buildFont(rPr, scaledPxPerEmu, themeData, defSz = 1800) {
  return buildFontInherited(rPr, null, scaledPxPerEmu, themeData, defSz);
}

// ── Registry inspection ──────────────────────────────────────────────────────

/**
 * List all currently registered custom fonts.
 * @returns {{ family: string, weight: string, style: string }[]}
 */
export function listRegisteredFonts() {
  const result = [];
  for (const [family, faces] of _customFonts) {
    for (const face of faces) {
      result.push({ family, weight: face.weight, style: face.style, status: face.status });
    }
  }
  return result;
}

/**
 * Remove all registered custom fonts.
 * Useful when re-using the renderer with a different brand kit.
 */
export function clearRegisteredFonts() {
  for (const [, faces] of _customFonts) {
    for (const face of faces) {
      try { document.fonts.delete(face); } catch (_) {}
    }
  }
  _customFonts.clear();
  // Remove custom overrides from MS_FONT_MAP
  for (const [key, val] of Object.entries(MS_FONT_MAP)) {
    if (val._custom) delete MS_FONT_MAP[key];
  }
  // Clear system font cache entries that were custom
  // (they'll be re-detected next time)
}
