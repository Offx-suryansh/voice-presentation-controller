/**
 * zip.js — Lightweight ZIP/ZIP64 reader, zero dependencies.
 *
 * Uses the browser/Node built-in DecompressionStream API for DEFLATE.
 * Browser support: Chrome 80+, Firefox 113+, Safari 16.4+, Node 18+
 *
 * Supported compression methods:
 *   0  = STORED   (no compression)
 *   8  = DEFLATE  (standard)
 */

/**
 * Parse a ZIP archive.
 * @param {ArrayBuffer|Uint8Array} input
 * @returns {Promise<Record<string, Uint8Array>>} path → bytes map
 */
export async function readZip(input) {
  const data = input instanceof Uint8Array ? input : new Uint8Array(input);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);

  // ── Locate EOCD (End of Central Directory) ──────────────────────────────────
  // Signature 0x06054b50.  The comment field can be up to 65535 bytes, so we
  // search backwards that far.  Stop as soon as we find a valid signature that
  // yields a sane central-directory offset.
  const EOCD_SIG   = 0x06054b50;
  const EOCD64_SIG = 0x06064b50; // ZIP64 end record
  const EOCD64_LOC = 0x07064b50; // ZIP64 end locator

  let eocdOff = -1;
  const searchStart = Math.max(0, data.length - 65557);
  for (let i = data.length - 22; i >= searchStart; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocdOff = i;
      break;
    }
  }
  if (eocdOff === -1) throw new Error('Not a valid ZIP file: EOCD not found');

  // ── Check for ZIP64 ─────────────────────────────────────────────────────────
  // If the EOCD64 locator sits just before the EOCD, use the ZIP64 end record.
  let cdOffset, cdCount;

  const locatorOff = eocdOff - 20;
  if (locatorOff >= 0 && view.getUint32(locatorOff, true) === EOCD64_LOC) {
    // ZIP64: offset of EOCD64 record is an 8-byte value at locator+8
    const eocd64Off = Number(view.getBigUint64(locatorOff + 8, true));
    if (view.getUint32(eocd64Off, true) !== EOCD64_SIG) {
      throw new Error('Invalid ZIP64 end record');
    }
    cdOffset = Number(view.getBigUint64(eocd64Off + 48, true));
    cdCount  = Number(view.getBigUint64(eocd64Off + 32, true));
  } else {
    // Standard ZIP32
    cdOffset = view.getUint32(eocdOff + 16, true);
    cdCount  = view.getUint16(eocdOff +  8, true);
  }

  if (cdOffset + 4 > data.length) {
    throw new Error('ZIP central directory offset is outside the file');
  }

  // ── Walk Central Directory ──────────────────────────────────────────────────
  const CD_SIG = 0x02014b50;
  const files  = {};
  let pos = cdOffset;

  for (let i = 0; i < cdCount; i++) {
    if (pos + 46 > data.length) break;
    if (view.getUint32(pos, true) !== CD_SIG) break;

    const method      = view.getUint16(pos + 10, true);
    let   compSize    = view.getUint32(pos + 20, true);
    let   uncompSize  = view.getUint32(pos + 24, true);
    const nameLen     = view.getUint16(pos + 28, true);
    const extraLen    = view.getUint16(pos + 30, true);
    const commentLen  = view.getUint16(pos + 32, true);
    let   localOffset = view.getUint32(pos + 42, true);
    const name        = utf8(data, pos + 46, nameLen);

    // ── ZIP64 extra field ────────────────────────────────────────────────────
    // If any 32-bit size fields are 0xFFFFFFFF, read true values from extra.
    if (compSize === 0xFFFFFFFF || uncompSize === 0xFFFFFFFF || localOffset === 0xFFFFFFFF) {
      const extraStart = pos + 46 + nameLen;
      const extraEnd   = extraStart + extraLen;
      let ep = extraStart;
      while (ep + 4 <= extraEnd) {
        const tag  = view.getUint16(ep,     true);
        const size = view.getUint16(ep + 2, true);
        if (tag === 0x0001) { // ZIP64 extended info
          let off = ep + 4;
          if (uncompSize  === 0xFFFFFFFF && off + 8 <= extraEnd) { uncompSize  = Number(view.getBigUint64(off, true)); off += 8; }
          if (compSize    === 0xFFFFFFFF && off + 8 <= extraEnd) { compSize    = Number(view.getBigUint64(off, true)); off += 8; }
          if (localOffset === 0xFFFFFFFF && off + 8 <= extraEnd) { localOffset = Number(view.getBigUint64(off, true)); }
          break;
        }
        ep += 4 + size;
      }
    }

    pos += 46 + nameLen + extraLen + commentLen;

    if (name.endsWith('/')) continue; // directory entry — skip
    if (method !== 0 && method !== 8) continue; // unsupported compression

    // ── Local File Header ────────────────────────────────────────────────────
    if (localOffset + 30 > data.length) continue;
    const lhNameLen  = view.getUint16(localOffset + 26, true);
    const lhExtraLen = view.getUint16(localOffset + 28, true);
    const dataStart  = localOffset + 30 + lhNameLen + lhExtraLen;
    if (dataStart + compSize > data.length) continue;

    const compData = data.subarray(dataStart, dataStart + compSize);

    try {
      files[name] = method === 0
        ? compData.slice()                      // STORED
        : await inflateRaw(compData);           // DEFLATE
    } catch (e) {
      console.warn(`[zip.js] Failed to decompress "${name}":`, e.message);
    }
  }

  return files;
}

// ── Decompression ─────────────────────────────────────────────────────────────

/**
 * Decompress raw DEFLATE bytes using the native DecompressionStream API.
 * Much faster than a JS implementation and ~0 bundle cost.
 */
async function inflateRaw(compData) {
  const ds     = new DecompressionStream('deflate-raw');
  const writer = ds.writable.getWriter();
  const reader = ds.readable.getReader();

  writer.write(compData);
  writer.close();

  const chunks = [];
  let totalLen = 0;

  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    chunks.push(value);
    totalLen += value.length;
  }

  // Concatenate all output chunks into a single typed array
  const out = new Uint8Array(totalLen);
  let off = 0;
  for (const c of chunks) { out.set(c, off); off += c.length; }
  return out;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Decode a UTF-8 slice of a Uint8Array. */
function utf8(data, start, len) {
  return new TextDecoder().decode(data.subarray(start, start + len));
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Read a file from the ZIP as a UTF-8 string.
 * @param {Record<string,Uint8Array>} files
 * @param {string} path
 * @returns {string|null}
 */
export function getFileText(files, path) {
  const d = files[path];
  return d ? new TextDecoder().decode(d) : null;
}

/**
 * Get raw bytes for a file in the ZIP.
 * @param {Record<string,Uint8Array>} files
 * @param {string} path
 * @returns {Uint8Array|null}
 */
export function getFileBytes(files, path) {
  return files[path] ?? null;
}

/**
 * List all file paths in the ZIP (excluding directories).
 * @param {Record<string,Uint8Array>} files
 * @returns {string[]}
 */
export function listFiles(files) {
  return Object.keys(files);
}
