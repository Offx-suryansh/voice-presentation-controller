/**
 * fntdata.js — Embedded PPTX font decoder.
 *
 * PowerPoint embeds fonts as .fntdata files inside ppt/fonts/.
 * These are standard TrueType/OpenType fonts XOR-obfuscated with a key
 * derived from the relationship ID, as specified in ECMA-376 §15.2.12.
 *
 * The 32-byte XOR key is constructed from the relationship GUID and applied
 * to the first 32 bytes of the font data. The rest of the file is plain TTF.
 *
 * After decoding, the font is loaded via the FontFace API so it becomes
 * available for canvas text rendering automatically.
 *
 * Usage:
 *   await loadEmbeddedFonts(renderer._files, renderer._allRels);
 *   // fonts are now available — renderer will use them automatically
 *
 * Or via the renderer instance:
 *   await renderer.loadEmbeddedFonts();
 */

// ── XOR key derivation (ECMA-376 §15.2.12) ────────────────────────────────────

/**
 * Derive the 32-byte obfuscation key from a relationship ID string.
 *
 * The relationship ID is a GUID in the form:
 *   {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
 *
 * The key is constructed by parsing the GUID into its binary representation
 * and XOR-interleaving according to the spec.
 *
 * @param {string} rId — relationship GUID (from fontRef in presentation.xml)
 * @returns {Uint8Array} 32-byte key
 */
function deriveObfuscationKey(rId) {
  // Extract hex digits from GUID, stripping braces and dashes
  const hex = rId.replace(/[{}\-]/g, '');
  if (hex.length < 32) {
    // Pad or repeat if too short (non-standard GUID)
    const padded = hex.padEnd(32, '0');
    return hexToBytes(padded.slice(0, 32));
  }

  // ECMA-376 specifies the key as the GUID bytes in a specific byte order:
  // Data1 (4 bytes, little-endian), Data2 (2 bytes, little-endian),
  // Data3 (2 bytes, little-endian), Data4 (8 bytes, big-endian)
  const data1 = hex.slice(0, 8);    // 4 bytes
  const data2 = hex.slice(8, 12);   // 2 bytes
  const data3 = hex.slice(12, 16);  // 2 bytes
  const data4 = hex.slice(16, 32);  // 8 bytes

  const key = new Uint8Array(16);
  // Data1: little-endian
  key[0] = parseInt(data1.slice(6, 8), 16);
  key[1] = parseInt(data1.slice(4, 6), 16);
  key[2] = parseInt(data1.slice(2, 4), 16);
  key[3] = parseInt(data1.slice(0, 2), 16);
  // Data2: little-endian
  key[4] = parseInt(data2.slice(2, 4), 16);
  key[5] = parseInt(data2.slice(0, 2), 16);
  // Data3: little-endian
  key[6] = parseInt(data3.slice(2, 4), 16);
  key[7] = parseInt(data3.slice(0, 2), 16);
  // Data4: big-endian
  for (let i = 0; i < 8; i++) {
    key[8 + i] = parseInt(data4.slice(i * 2, i * 2 + 2), 16);
  }

  // The spec says the 32-byte key is formed by repeating the 16-byte GUID key twice
  const key32 = new Uint8Array(32);
  key32.set(key);
  key32.set(key, 16);
  return key32;
}

function hexToBytes(hex) {
  const bytes = new Uint8Array(hex.length / 2);
  for (let i = 0; i < bytes.length; i++) {
    bytes[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  }
  return bytes;
}

// ── Decode font data ──────────────────────────────────────────────────────────

/**
 * Decode an .fntdata buffer by XOR-ing the first 32 bytes with the key.
 * Returns a Uint8Array containing valid TTF/OTF font data.
 *
 * @param {Uint8Array} data — raw .fntdata file contents
 * @param {Uint8Array} key  — 32-byte XOR key from deriveObfuscationKey()
 * @returns {Uint8Array}    — decoded font bytes
 */
function decodeFontData(data, key) {
  const decoded = new Uint8Array(data);
  // XOR only the first 32 bytes
  for (let i = 0; i < Math.min(32, decoded.length); i++) {
    decoded[i] ^= key[i];
  }
  return decoded;
}

/**
 * Detect if decoded bytes look like a valid font.
 * TTF starts with 0x00010000, OTF with 'OTTO', woff with 'wOFF'.
 */
function isValidFont(bytes) {
  if (bytes.length < 4) return false;
  const sig = (bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3];
  return sig === 0x00010000   // TTF
    || sig === 0x4F54544F     // 'OTTO' - OTF CFF
    || sig === 0x74727565     // 'true'
    || sig === 0x774F4646     // 'wOFF'
    || sig === 0x774F4632;    // 'wOF2'
}

// ── Font registry ─────────────────────────────────────────────────────────────

/** Tracks which embedded fonts have been loaded to avoid double-loading. */
const _loadedEmbedded = new Set();

// ── Main decoder ──────────────────────────────────────────────────────────────

/**
 * Parse presentation.xml's <p:embeddedFontLst> to find embedded fonts,
 * decode them, and load them via FontFace API.
 *
 * @param {Record<string,Uint8Array>} files  — ZIP file map
 * @param {object} presRels                  — rels for presentation.xml
 * @returns {Promise<EmbeddedFontResult[]>}
 */
export async function loadEmbeddedFonts(files, presRels) {
  const results = [];

  // Find presentation.xml
  const presXml = files['ppt/presentation.xml'];
  if (!presXml) return results;

  const parser = new DOMParser();
  const presDoc = parser.parseFromString(new TextDecoder().decode(presXml), 'application/xml');

  const embeddedFontLst = presDoc.querySelector('embeddedFontLst') ||
    [...presDoc.getElementsByTagName('*')].find(el => el.localName === 'embeddedFontLst');

  if (!embeddedFontLst) return results;

  // Each <p:embeddedFont> has a <p:font typeface="..."/> and <p:regular/bold/italic/boldItalic r:id="rId..."/>
  const embeddedFonts = [...embeddedFontLst.children].filter(el => el.localName === 'embeddedFont');

  for (const fontEl of embeddedFonts) {
    const fontDescEl = [...fontEl.children].find(el => el.localName === 'font');
    if (!fontDescEl) continue;

    const typeface = fontDescEl.getAttribute('typeface') || fontDescEl.getAttribute('t');
    if (!typeface) continue;

    const variants = [
      { el: [...fontEl.children].find(e => e.localName === 'regular'),   weight: '400', style: 'normal' },
      { el: [...fontEl.children].find(e => e.localName === 'bold'),      weight: '700', style: 'normal' },
      { el: [...fontEl.children].find(e => e.localName === 'italic'),    weight: '400', style: 'italic' },
      { el: [...fontEl.children].find(e => e.localName === 'boldItalic'),weight: '700', style: 'italic' },
    ];

    for (const { el, weight, style } of variants) {
      if (!el) continue;

      // Get rId — attribute may be r:id or just id
      const rId = el.getAttribute('r:id') || el.getAttribute('id') || el.getAttributeNS(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'
      );
      if (!rId) continue;

      const cacheKey = `${typeface}:${weight}:${style}`;
      if (_loadedEmbedded.has(cacheKey)) {
        results.push({ typeface, weight, style, status: 'already-loaded' });
        continue;
      }

      // Resolve relationship to file path
      const rel = presRels ? presRels[rId] : null;
      if (!rel) {
        results.push({ typeface, weight, style, status: 'rel-not-found', rId });
        continue;
      }

      const rawData = files[rel.fullPath];
      if (!rawData) {
        results.push({ typeface, weight, style, status: 'file-not-found', path: rel.fullPath });
        continue;
      }

      // Derive decryption key from the rId
      // The rId used for key derivation is the GUID from the relationship target, not the Id
      // In practice many tools use the rId string directly; we try both approaches.
      let decoded = null;
      const key = deriveObfuscationKey(rId);
      const attempt1 = decodeFontData(rawData, key);

      if (isValidFont(attempt1)) {
        decoded = attempt1;
      } else {
        // Try with the full target path's basename as GUID
        const basename = rel.fullPath.split('/').pop().replace('.fntdata', '');
        const key2 = deriveObfuscationKey(basename);
        const attempt2 = decodeFontData(rawData, key2);
        if (isValidFont(attempt2)) {
          decoded = attempt2;
        } else {
          // Try raw (some .fntdata files aren't actually obfuscated)
          if (isValidFont(rawData)) {
            decoded = rawData;
          }
        }
      }

      if (!decoded) {
        results.push({ typeface, weight, style, status: 'decode-failed', path: rel.fullPath });
        continue;
      }

      // Load via FontFace API
      try {
        const fontFace = new FontFace(typeface, decoded.buffer, { weight, style });
        await fontFace.load();
        document.fonts.add(fontFace);
        _loadedEmbedded.add(cacheKey);
        results.push({ typeface, weight, style, status: 'loaded', path: rel.fullPath });
      } catch (err) {
        results.push({ typeface, weight, style, status: 'load-failed', error: err.message });
      }
    }
  }

  return results;
}

/**
 * Get info about embedded fonts without loading them.
 * Useful for displaying what fonts are embedded in a PPTX.
 *
 * @param {Record<string,Uint8Array>} files
 * @returns {EmbeddedFontInfo[]}
 */
export function listEmbeddedFonts(files) {
  const presXml = files['ppt/presentation.xml'];
  if (!presXml) return [];

  const presDoc = new DOMParser().parseFromString(
    new TextDecoder().decode(presXml), 'application/xml'
  );

  const embeddedFonts = [...presDoc.getElementsByTagName('*')]
    .filter(el => el.localName === 'embeddedFont');

  return embeddedFonts.map(fontEl => {
    const fontDescEl = [...fontEl.children].find(el => el.localName === 'font');
    const typeface = fontDescEl ? (fontDescEl.getAttribute('typeface') || fontDescEl.getAttribute('t') || '') : '';
    const variants = ['regular', 'bold', 'italic', 'boldItalic']
      .filter(v => [...fontEl.children].some(el => el.localName === v));
    return { typeface, variants, loaded: variants.some(v => _loadedEmbedded.has(
      `${typeface}:${v === 'bold' || v === 'boldItalic' ? '700' : '400'}:${v.includes('talic') ? 'italic' : 'normal'}`
    ))};
  }).filter(f => f.typeface);
}
