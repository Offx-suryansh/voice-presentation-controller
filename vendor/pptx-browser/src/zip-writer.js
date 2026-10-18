/**
 * zip-writer.js — ZIP archive serializer. Zero dependencies.
 *
 * Uses the native CompressionStream('deflate-raw') API for compression.
 * Browser support: Chrome 80+, Firefox 113+, Safari 16.4+, Node 18+
 * (same as the existing zip.js reader).
 *
 * Usage:
 *   const w = new ZipWriter();
 *   w.addText('hello.txt', 'Hello world');
 *   w.addBytes('data.bin', uint8array);
 *   const zipBytes = await w.finalize();  // → Uint8Array
 */

// ── DEFLATE via CompressionStream ─────────────────────────────────────────────

async function deflateRaw(data) {
  const cs     = new CompressionStream('deflate-raw');
  const writer = cs.writable.getWriter();
  const reader = cs.readable.getReader();

  writer.write(data);
  writer.close();

  const chunks = [];
  let total = 0;
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.length;
  }
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) { out.set(c, off); off += c.length; }
  return out;
}

// ── CRC-32 ────────────────────────────────────────────────────────────────────

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();

function crc32(data) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) crc = CRC_TABLE[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// ── Binary helpers ────────────────────────────────────────────────────────────

class BufWriter {
  constructor() { this._chunks = []; this._size = 0; }
  append(u8) { this._chunks.push(u8); this._size += u8.length; }
  get size() { return this._size; }
  concat() {
    const out = new Uint8Array(this._size);
    let off = 0;
    for (const c of this._chunks) { out.set(c, off); off += c.length; }
    return out;
  }
}

function u16le(n) { return new Uint8Array([(n) & 0xFF, (n >> 8) & 0xFF]); }
function u32le(n) {
  return new Uint8Array([(n) & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >> 24) & 0xFF]);
}

const enc = new TextEncoder();
function utf8bytes(s) { return enc.encode(s); }

// MS-DOS epoch: 1980-01-01 00:00:00
const DOS_DATE = 0x4A21; // 2025-01-01
const DOS_TIME = 0x0000;

// ── ZipWriter ────────────────────────────────────────────────────────────────

export class ZipWriter {
  constructor() {
    /** @type {Array<{name, nameBytes, compData, uncompSize, crc, method, localOffset}>} */
    this._entries = [];
  }

  /**
   * Add a file from a UTF-8 string.
   * @param {string} name
   * @param {string} text
   */
  async addText(name, text) {
    return this.addBytes(name, enc.encode(text));
  }

  /**
   * Add a file from raw bytes, with optional DEFLATE compression.
   * @param {string}     name
   * @param {Uint8Array} data
   * @param {boolean}    [compress=true]  false for already-compressed data
   */
  async addBytes(name, data, compress = true) {
    const nameBytes  = utf8bytes(name);
    const crc        = crc32(data);
    const uncompSize = data.length;

    let compData, method;
    if (compress && data.length > 32) {
      const deflated = await deflateRaw(data);
      // Only use deflated if it's actually smaller
      if (deflated.length < data.length) {
        compData = deflated;
        method   = 8; // DEFLATE
      } else {
        compData = data;
        method   = 0; // STORED
      }
    } else {
      compData = data;
      method   = 0; // STORED
    }

    this._entries.push({ name, nameBytes, compData, uncompSize, crc, method, localOffset: 0 });
  }

  /**
   * Serialize the archive and return the ZIP bytes.
   * @returns {Promise<Uint8Array>}
   */
  async finalize() {
    const body = new BufWriter();
    const cdEntries = [];

    // ── Local file headers + data ─────────────────────────────────────────
    for (const entry of this._entries) {
      entry.localOffset = body.size;

      // Local file header  (signature 0x04034b50)
      body.append(new Uint8Array([0x50, 0x4B, 0x03, 0x04])); // signature
      body.append(u16le(20));              // version needed: 2.0
      body.append(u16le(0x800));           // general purpose bit flag: UTF-8 name
      body.append(u16le(entry.method));    // compression method
      body.append(u16le(DOS_TIME));        // last mod time
      body.append(u16le(DOS_DATE));        // last mod date
      body.append(u32le(entry.crc));       // CRC-32
      body.append(u32le(entry.compData.length));   // compressed size
      body.append(u32le(entry.uncompSize));         // uncompressed size
      body.append(u16le(entry.nameBytes.length));  // filename length
      body.append(u16le(0));               // extra field length
      body.append(entry.nameBytes);
      body.append(entry.compData);
    }

    const cdOffset = body.size;

    // ── Central directory ─────────────────────────────────────────────────
    for (const entry of this._entries) {
      body.append(new Uint8Array([0x50, 0x4B, 0x01, 0x02])); // CD signature
      body.append(u16le(0x031E));           // version made by: Unix, 3.0
      body.append(u16le(20));              // version needed
      body.append(u16le(0x800));           // general purpose bit flag
      body.append(u16le(entry.method));
      body.append(u16le(DOS_TIME));
      body.append(u16le(DOS_DATE));
      body.append(u32le(entry.crc));
      body.append(u32le(entry.compData.length));
      body.append(u32le(entry.uncompSize));
      body.append(u16le(entry.nameBytes.length));
      body.append(u16le(0));               // extra field length
      body.append(u16le(0));               // file comment length
      body.append(u16le(0));               // disk number start
      body.append(u16le(0));               // internal file attributes
      body.append(u32le(0));               // external file attributes
      body.append(u32le(entry.localOffset));
      body.append(entry.nameBytes);
      cdEntries.push(entry);
    }

    const cdSize = body.size - cdOffset;

    // ── End of central directory ──────────────────────────────────────────
    body.append(new Uint8Array([0x50, 0x4B, 0x05, 0x06])); // EOCD signature
    body.append(u16le(0));                          // disk number
    body.append(u16le(0));                          // start disk
    body.append(u16le(this._entries.length));       // entries on this disk
    body.append(u16le(this._entries.length));       // total entries
    body.append(u32le(cdSize));                     // central dir size
    body.append(u32le(cdOffset));                   // central dir offset
    body.append(u16le(0));                          // comment length

    return body.concat();
  }
}

/**
 * Convenience: create a ZIP from a files map (path → Uint8Array).
 * @param {Record<string, Uint8Array|string>} files
 * @returns {Promise<Uint8Array>}
 */
export async function writeZip(files) {
  const w = new ZipWriter();
  for (const [path, data] of Object.entries(files)) {
    if (typeof data === 'string') {
      await w.addText(path, data);
    } else {
      await w.addBytes(path, data);
    }
  }
  return w.finalize();
}
