/**
 * pdf.js — PPTX → PDF exporter. Zero dependencies.
 *
 * Renders each slide onto a canvas at target resolution, encodes as JPEG,
 * and packages into a valid PDF/1.4 binary that any PDF viewer can open.
 *
 * Output is pixel-perfect — same rendering as the canvas renderer, just
 * wrapped in a PDF container.
 *
 * Usage:
 *   import { exportToPdf } from 'pptx-canvas-renderer';
 *
 *   // Render all slides → PDF bytes
 *   const pdfBytes = await exportToPdf(renderer);
 *
 *   // Download as file
 *   await downloadAsPdf(renderer, 'presentation.pdf');
 *
 *   // Or via renderer instance:
 *   const bytes = await renderer.toPdf({ quality: 0.92, width: 1920 });
 *   await renderer.downloadPdf('presentation.pdf');
 *
 * Options:
 *   width    {number}  render width per slide in pixels  (default 1920)
 *   quality  {number}  JPEG quality 0..1                 (default 0.92)
 *   slides   {number[]} only export specified slides
 *   onProgress {function} (done, total) => void
 */

// ── PDF binary helpers ────────────────────────────────────────────────────────

const CR = 0x0D, LF = 0x0A;
const enc = new TextEncoder();

// ── DPI → pixel width conversion ─────────────────────────────────────────────

/**
 * Convert a DPI value to a pixel width for the given renderer's slide size.
 * slideSize.cx is in EMU (914400 EMU = 1 inch).
 */
function dpiToWidth(renderer, dpi) {
  const inches = renderer.slideSize.cx / 914400;
  return Math.round(inches * dpi);
}

class PdfBuf {
  constructor() { this._parts = []; this._size = 0; }
  write(s) {
    const b = typeof s === 'string' ? enc.encode(s) : s;
    this._parts.push(b);
    this._size += b.length;
    return this._size; // returns position AFTER this write
  }
  get size() { return this._size; }
  concat() {
    const out = new Uint8Array(this._size);
    let off = 0;
    for (const p of this._parts) { out.set(p, off); off += p.length; }
    return out;
  }
}

/** Extract raw JPEG bytes from a canvas element. Returns Uint8Array. */
async function canvasToJpeg(canvas, quality = 0.92) {
  // OffscreenCanvas uses convertToBlob(); HTMLCanvasElement uses toBlob()
  if (typeof canvas.convertToBlob === 'function') {
    const blob = await canvas.convertToBlob({ type: 'image/jpeg', quality });
    const buf = await blob.arrayBuffer();
    return new Uint8Array(buf);
  }
  return new Promise((resolve, reject) => {
    canvas.toBlob(blob => {
      if (!blob) { reject(new Error('canvas.toBlob failed')); return; }
      blob.arrayBuffer().then(buf => resolve(new Uint8Array(buf))).catch(reject);
    }, 'image/jpeg', quality);
  });
}

/** Render a slide to an HTMLCanvasElement at given pixel width. */
async function renderSlide(renderer, slideIndex, widthPx) {
  const { cx, cy } = renderer.slideSize;
  const h = Math.round(widthPx * cy / cx);
  const canvas = typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(widthPx, h)
    : Object.assign(document.createElement('canvas'), { width: widthPx, height: h });
  await renderer.renderSlide(slideIndex, canvas, widthPx);
  return canvas;
}

// ── PDF structure constants ───────────────────────────────────────────────────

// All measurements in PDF points (1pt = 1/72 inch)
// Standard slide = 10 inches wide × 7.5 inches tall = 720pt × 540pt
// Widescreen (16:9) = 13.33 × 7.5 inches

function slidePageSize(cx, cy) {
  // Convert EMU to points: 1 inch = 914400 EMU = 72 pt
  const w = (cx / 914400) * 72;
  const h = (cy / 914400) * 72;
  return { w, h };
}

// ── PDF object writer ─────────────────────────────────────────────────────────

class PdfWriter {
  constructor() {
    this._buf    = new PdfBuf();
    this._xref   = [];   // byte offsets of each object
    this._objNum = 0;    // current object counter (0 = free entry)
  }

  _nextObj() { return ++this._objNum; }

  /** Write a PDF object. Returns its object number. */
  _writeObj(num, dictStr, streamData = null) {
    const off = this._buf.size;
    this._xref[num] = off;

    this._buf.write(`${num} 0 obj\n`);
    if (streamData) {
      const len = streamData.length;
      this._buf.write(`${dictStr.replace('__LEN__', len)}\n`);
      this._buf.write('stream\r\n');
      this._buf.write(streamData);
      this._buf.write('\r\nendstream\n');
    } else {
      this._buf.write(`${dictStr}\n`);
    }
    this._buf.write('endobj\n\n');
    return num;
  }

  /** Reserve the next N object numbers. Returns first number. */
  _reserveObjs(n) {
    const first = this._objNum + 1;
    this._objNum += n;
    return first;
  }

  // ── Build PDF ───────────────────────────────────────────────────────────────

  async build(renderer, opts = {}) {
    const {
      widthPx    = 1920,
      quality    = 0.92,
      slideList  = null,
      onProgress = null,
    } = opts;

    const indices = slideList ?? Array.from({ length: renderer.slideCount }, (_, i) => i);
    const n       = indices.length;

    const { cx, cy }  = renderer.slideSize;
    const { w: pgW, h: pgH } = slidePageSize(cx, cy);

    // ── PDF header ────────────────────────────────────────────────────────
    this._buf.write('%PDF-1.4\n');
    // Comment with 4 high bytes signals binary content to transfer tools
    this._buf.write('%\xFF\xFE\xFD\xFC\n\n');

    // ── Pre-allocate object numbers ───────────────────────────────────────
    // Object layout:
    //   1 = Catalog
    //   2 = Pages
    //   3..3+n-1      = Page objects
    //   3+n..3+2n-1   = Image XObjects (one per slide)
    //   3+2n..3+3n-1  = Content streams (one per slide)

    const catalogNum  = this._nextObj();   // 1
    const pagesNum    = this._nextObj();   // 2
    const pageNums    = Array.from({ length: n }, () => this._nextObj());
    const imageNums   = Array.from({ length: n }, () => this._nextObj());
    const contentNums = Array.from({ length: n }, () => this._nextObj());

    // ── Catalog ───────────────────────────────────────────────────────────
    this._writeObj(catalogNum, `<< /Type /Catalog /Pages ${pagesNum} 0 R >>`);

    // ── Pages (written later after we know all page refs) ─────────────────
    // We'll write pages dict here with placeholders and rewrite isn't easy
    // in streaming mode — instead we write it now knowing all page nums
    const kidsStr = pageNums.map(n => `${n} 0 R`).join(' ');
    this._writeObj(pagesNum, `<< /Type /Pages /Kids [${kidsStr}] /Count ${n} >>`);

    // ── Render each slide and write objects ───────────────────────────────
    for (let i = 0; i < n; i++) {
      const slideIdx = indices[i];
      onProgress?.(i, n);

      // Render
      const canvas  = await renderSlide(renderer, slideIdx, widthPx);
      const jpegData = await canvasToJpeg(canvas, quality);
      const imgW     = canvas.width;
      const imgH     = canvas.height;

      // Page object
      this._writeObj(pageNums[i],
        `<< /Type /Page /Parent ${pagesNum} 0 R ` +
        `/MediaBox [0 0 ${pgW.toFixed(3)} ${pgH.toFixed(3)}] ` +
        `/Resources << /XObject << /Im${i} ${imageNums[i]} 0 R >> >> ` +
        `/Contents ${contentNums[i]} 0 R >>`
      );

      // Image XObject
      this._writeObj(imageNums[i],
        `<< /Type /XObject /Subtype /Image ` +
        `/Width ${imgW} /Height ${imgH} ` +
        `/ColorSpace /DeviceRGB /BitsPerComponent 8 ` +
        `/Filter /DCTDecode /Length __LEN__ >>`,
        jpegData
      );

      // Content stream: place image to fill page
      const contentStr =
        `q ${pgW.toFixed(3)} 0 0 ${pgH.toFixed(3)} 0 0 cm /Im${i} Do Q`;
      this._writeObj(contentNums[i],
        `<< /Length __LEN__ >>`,
        enc.encode(contentStr)
      );
    }

    onProgress?.(n, n);

    // ── Cross-reference table ─────────────────────────────────────────────
    const xrefOffset = this._buf.size;
    const totalObjs  = this._objNum + 1; // +1 for the free entry at 0

    this._buf.write(`xref\n0 ${totalObjs}\n`);
    // Entry 0: free list head
    this._buf.write('0000000000 65535 f \r\n');
    for (let i = 1; i < totalObjs; i++) {
      const off = this._xref[i] ?? 0;
      this._buf.write(String(off).padStart(10, '0') + ' 00000 n \r\n');
    }

    // ── Trailer ───────────────────────────────────────────────────────────
    this._buf.write(`trailer\n<< /Size ${totalObjs} /Root ${catalogNum} 0 R >>\n`);
    this._buf.write(`startxref\n${xrefOffset}\n%%EOF\n`);

    return this._buf.concat();
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Export all (or selected) slides to a PDF binary.
 *
 * @param {object}   renderer        — loaded PptxRenderer instance
 * @param {object}   [opts]
 * @param {number}   [opts.width]           render width in pixels (overrides dpi)
 * @param {number}   [opts.dpi=150]          dots per inch (standard screen=96, print=300)
 * @param {number}   [opts.quality=0.92]   JPEG quality 0..1
 * @param {number[]} [opts.slides]         slide indices to include (default: all)
 * @param {function} [opts.onProgress]     (done, total) => void
 * @returns {Promise<Uint8Array>}  raw PDF bytes
 */
export async function exportToPdf(renderer, opts = {}) {
  const {
    width      = null,
    dpi        = 150,
    quality    = 0.92,
    slides     = null,
    onProgress = null,
  } = opts;

  const resolvedWidth = width ?? dpiToWidth(renderer, dpi);
  const writer = new PdfWriter();
  return writer.build(renderer, {
    widthPx:    resolvedWidth,
    quality,
    slideList:  slides,
    onProgress,
  });
}

/**
 * Export to PDF and trigger a browser download.
 *
 * @param {object}   renderer
 * @param {string}   [filename='presentation.pdf']
 * @param {object}   [opts]   — same as exportToPdf options (dpi, quality, slides…)
 * @returns {Promise<void>}
 */
export async function downloadAsPdf(renderer, filename = 'presentation.pdf', opts = {}) {
  const bytes = await exportToPdf(renderer, opts);
  const blob  = new Blob([bytes], { type: 'application/pdf' });
  const url   = URL.createObjectURL(blob);
  const a     = document.createElement('a');
  a.href     = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

/**
 * Export a single slide to PDF bytes.
 * @param {number} slideIndex
 * @param {object} renderer
 * @param {object} [opts]
 * @returns {Promise<Uint8Array>}
 */
export async function exportSlideToPdf(slideIndex, renderer, opts = {}) {
  return exportToPdf(renderer, { ...opts, slides: [slideIndex] });
}
