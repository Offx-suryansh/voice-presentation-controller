/**
 * index.js — PptxRenderer
 *
 * Load and render PPTX files onto an HTML Canvas element.
 * Zero dependencies — uses native browser APIs only.
 *
 * Usage
 * ─────
 *   import { PptxRenderer } from 'pptx-browser';
 *
 *   const renderer = new PptxRenderer();
 *   await renderer.load(fileOrArrayBuffer);          // load the PPTX
 *   console.log(renderer.slideCount);               // e.g. 12
 *   await renderer.renderSlide(0, canvasElement);   // render slide 0
 *   renderer.destroy();                             // free blob URLs
 *
 * Compatible environments
 * ────────────────────────
 *   Chrome 80+, Firefox 113+, Safari 16.4+, Node.js 18+ (with node-canvas)
 */

import { readZip } from './zip.js';
import { parseXml, g1, gtn, attr, attrInt } from './utils.js';
import { parseTheme, parseClrMap, buildThemeColors } from './theme.js';
import {
  collectUsedFonts, loadGoogleFontsFor,
  registerFont, registerFonts,
  detectEmbeddedFonts, listRegisteredFonts,
} from './fonts.js';
import {
  getRels, loadImages,
  renderBackground,
  renderShape, renderPicture, renderGroupShape,
  renderGraphicFrame, renderConnector,
  renderSpTree, buildPlaceholderMap,
} from './render.js';
import { PptxWriter } from './writer.js';
import { exportToPdf, downloadAsPdf, exportSlideToPdf } from './pdf.js';
import { parseAnimations, parseTransition, PptxPlayer } from './animation.js';
import { renderSlideToSvg, renderAllSlidesToSvg } from './svg.js';
import { loadEmbeddedFonts, listEmbeddedFonts } from './fntdata.js';
import { extractSlide, extractAll, extractText, searchSlides } from './extract.js';
import { SlideShow } from './slideshow.js';
import { copySlideToClipboard, downloadSlide, downloadAllSlides, createLazyDeck } from './clipboard.js';

export default class PptxRenderer {
  constructor() {
    /** @type {Record<string, Uint8Array>} raw files extracted from the ZIP */
    this._files = {};
    this.slideSize  = { cx: 9144000, cy: 5143500 }; // EMU
    this.slidePaths = [];
    /** Total number of slides */
    this.slideCount = 0;

    this.themeData    = null;
    this.themeColors  = {};
    this.masterDoc    = null;
    this.masterRels   = {};
    this.masterImages = {};

    /** Blob URLs to revoke on destroy() */
    this._blobUrls = [];

    /**
     * Fonts embedded in the PPTX (detected during load).
     * @type {Array}
     */
    this.embeddedFonts = [];
  }

  // ── Loading ──────────────────────────────────────────────────────────────

  /**
   * Load a PPTX file.
   * @param {File|Blob|ArrayBuffer|Uint8Array} source
   * @param {(progress: number, message: string) => void} [onProgress]
   */
  async load(source, onProgress = () => {}) {
    // Normalise input to ArrayBuffer
    let buf;
    if (source instanceof Uint8Array) {
      buf = source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
    } else if (source instanceof ArrayBuffer) {
      buf = source;
    } else if (typeof source.arrayBuffer === 'function') {
      buf = await source.arrayBuffer(); // File / Blob
    } else {
      throw new TypeError('PptxRenderer.load(): expected File, Blob, ArrayBuffer, or Uint8Array');
    }

    onProgress(0.05, 'Decompressing…');
    this._files = await readZip(buf);

    // ── presentation.xml ─────────────────────────────────────────────────
    onProgress(0.1, 'Parsing presentation…');
    const presXml = this._readText('ppt/presentation.xml');
    if (!presXml) throw new Error('Invalid PPTX: missing ppt/presentation.xml');
    const presDoc = parseXml(presXml);

    const sldSz = g1(presDoc, 'sldSz');
    if (sldSz) {
      this.slideSize = {
        cx: attrInt(sldSz, 'cx', 9144000),
        cy: attrInt(sldSz, 'cy', 5143500),
      };
    }

    const presRels = await getRels(this._files, 'ppt/presentation.xml');

    // Detect embedded fonts (informational)
    this.embeddedFonts = detectEmbeddedFonts(presDoc, presRels);
    if (this.embeddedFonts.length > 0) {
      const names = this.embeddedFonts.map(f => f.family).join(', ');
      console.info('[pptx-browser] PPTX contains embedded fonts: ' + names + '. Use registerFont() to supply woff2/ttf versions.');
    }

    const sldIdLst = g1(presDoc, 'sldIdLst');
    if (sldIdLst) {
      for (const sldId of sldIdLst.children) {
        if (sldId.localName !== 'sldId') continue;
        const rId = sldId.getAttribute('r:id') || sldId.getAttribute('id');
        const rel = presRels[rId];
        if (rel) this.slidePaths.push(rel.fullPath);
      }
    }
    this.slideCount = this.slidePaths.length;

    // ── Theme ─────────────────────────────────────────────────────────────
    onProgress(0.2, 'Loading theme…');
    let themePath = Object.values(presRels).find(r => r.type?.includes('theme'))?.fullPath;
    if (!themePath) {
      const masterRel2 = Object.values(presRels).find(r => r.type?.includes('slideMaster'));
      if (masterRel2) {
        const mr2 = await getRels(this._files, masterRel2.fullPath);
        themePath = Object.values(mr2).find(r => r.type?.includes('theme'))?.fullPath;
      }
    }
    if (themePath) {
      const themeXml = this._readText(themePath);
      if (themeXml) this.themeData = parseTheme(parseXml(themeXml));
    }
    if (!this.themeData) {
      this.themeData = { colors: {}, majorFont: 'Calibri Light', minorFont: 'Calibri' };
    }
    this.themeColors = { ...this.themeData.colors };

    // ── Slide master ──────────────────────────────────────────────────────
    onProgress(0.3, 'Loading master…');
    const masterRel = Object.values(presRels).find(r => r.type?.includes('slideMaster'));
    if (masterRel) {
      const masterXml = this._readText(masterRel.fullPath);
      if (masterXml) {
        this.masterDoc    = parseXml(masterXml);
        this.masterRels   = await getRels(this._files, masterRel.fullPath);
        this.masterImages = await loadImages(this._files, this.masterRels);
        this._trackBlobs(this.masterImages);
        const clrMap = parseClrMap(this.masterDoc);
        this.themeColors = buildThemeColors(this.themeData, clrMap);
      }
    }

    onProgress(1, 'Ready');
  }

  // ── Rendering ─────────────────────────────────────────────────────────────

  /**
   * Render a single slide onto a canvas element.
   *
   * @param {number}            slideIndex  0-based slide index
   * @param {HTMLCanvasElement} canvas
   * @param {number}            [width=1280]  output canvas width in pixels
   */
  async renderSlide(slideIndex, canvas, width = 1280) {
    if (slideIndex < 0 || slideIndex >= this.slidePaths.length) {
      throw new RangeError(`Slide ${slideIndex} out of range (0–${this.slidePaths.length - 1})`);
    }

    const slidePath   = this.slidePaths[slideIndex];
    const slideXml    = this._readText(slidePath);
    if (!slideXml) throw new Error(`Could not read slide: ${slidePath}`);

    const slideDoc    = parseXml(slideXml);
    const slideRels   = await getRels(this._files, slidePath);
    const slideImages = await loadImages(this._files, slideRels);
    this._trackBlobs(slideImages);

    // Layout
    let layoutDoc = null, layoutRels = {}, layoutImages = {};
    const layoutRel = Object.values(slideRels).find(r => r.type?.includes('slideLayout'));
    if (layoutRel) {
      const layoutXml = this._readText(layoutRel.fullPath);
      if (layoutXml) {
        layoutDoc    = parseXml(layoutXml);
        layoutRels   = await getRels(this._files, layoutRel.fullPath);
        layoutImages = await loadImages(this._files, layoutRels);
        this._trackBlobs(layoutImages);
      }
    }

    const allImages = { ...this.masterImages, ...layoutImages, ...slideImages };
    const placeholderMap = buildPlaceholderMap([layoutDoc, this.masterDoc]);

    // Load Google Font substitutes for MS fonts used in this slide
    const usedFonts = collectUsedFonts([slideDoc, layoutDoc, this.masterDoc]);
    await loadGoogleFontsFor(usedFonts, this.themeData);

    // Canvas setup
    const scale  = width / this.slideSize.cx;
    const height = Math.round(this.slideSize.cy * scale);
    canvas.width  = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, width, height);

    // Background
    await renderBackground(
      ctx, slideDoc, this.masterDoc, layoutDoc,
      slideRels, this.masterRels,
      allImages, this.themeColors,
      scale, this.slideSize.cx, this.slideSize.cy,
    );

    // Master decorative shapes (non-placeholder)
    await this._renderNonPlaceholders(ctx, this.masterDoc, this.masterRels, this.masterImages, scale);

    // Layout decorative shapes (non-placeholder)
    await this._renderNonPlaceholders(ctx, layoutDoc, layoutRels, layoutImages, scale);

    // Slide content
    const cSld   = g1(slideDoc, 'cSld');
    const spTree = cSld ? g1(cSld, 'spTree') : null;
    if (spTree) {
      await renderSpTree(ctx, spTree, slideRels, allImages,
        this.themeColors, this.themeData, scale, placeholderMap, this._files);
    }
  }

  /**
   * Render all slides and return an array of canvas elements.
   * Useful for generating thumbnails.
   *
   * @param {number} [width=320]
   * @returns {Promise<HTMLCanvasElement[]>}
   */
  async renderAllSlides(width = 320) {
    const canvases = [];
    for (let i = 0; i < this.slideCount; i++) {
      const c = document.createElement('canvas');
      await this.renderSlide(i, c, width);
      canvases.push(c);
    }
    return canvases;
  }

  // ── Font management ────────────────────────────────────────────────────────

  /**
   * Register a custom font for this rendering session.
   * Call before renderSlide() for fonts that should be used in the PPTX.
   *
   * Accepts woff, woff2, ttf, otf files.
   *
   * @param {string} family
   *   The font name exactly as it appears in the PPTX, OR the MS Office name
   *   it should replace (e.g. "Calibri" to override the default substitute).
   * @param {string|URL|File|Blob|ArrayBuffer|Uint8Array} source
   * @param {object} [descriptors] — FontFace descriptors: { weight, style, unicodeRange }
   * @returns {Promise<FontFace>}
   *
   * @example
   * // Load regular weight
   * await renderer.registerFont('Acme Sans', '/fonts/acme-sans.woff2');
   *
   * // Load bold variant
   * await renderer.registerFont('Acme Sans', '/fonts/acme-sans-bold.woff2', { weight: '700' });
   *
   * // From a File object (e.g. dropped by user)
   * await renderer.registerFont('Acme Sans', file);
   *
   * // Override Calibri with your brand font
   * await renderer.registerFont('Calibri', '/fonts/brand.woff2');
   */
  registerFont(family, source, descriptors) {
    return registerFont(family, source, descriptors);
  }

  /**
   * Register multiple custom fonts at once.
   *
   * @param {Record<string, string | Array<{url: string, weight?: string, style?: string}>>} fontMap
   * @returns {Promise<void>}
   *
   * @example
   * await renderer.registerFonts({
   *   'Brand Sans': '/fonts/brand-sans.woff2',
   *   'Brand Serif': [
   *     { url: '/fonts/brand-serif.woff2',      weight: '400' },
   *     { url: '/fonts/brand-serif-bold.woff2', weight: '700' },
   *   ],
   * });
   */
  registerFonts(fontMap) {
    return registerFonts(fontMap);
  }

  /**
   * List all custom fonts currently registered.
   * @returns {{ family: string, weight: string, style: string, status: string }[]}
   */
  listRegisteredFonts() {
    return listRegisteredFonts();
  }

  // ── Private helpers ───────────────────────────────────────────────────────

  _readText(path) {
    const data = this._files[path];
    return data ? new TextDecoder().decode(data) : null;
  }

  _trackBlobs(imageCache) {
    for (const img of Object.values(imageCache)) {
      if (img?.src?.startsWith('blob:')) this._blobUrls.push(img.src);
    }
  }

  async _renderNonPlaceholders(ctx, doc, rels, images, scale) {
    if (!doc) return;
    const cSld   = g1(doc, 'cSld');
    const spTree = cSld ? g1(cSld, 'spTree') : null;
    if (!spTree) return;

    for (const child of spTree.children) {
      const ln = child.localName;
      if (!['sp','pic','grpSp','graphicFrame','cxnSp'].includes(ln)) continue;

      // Skip placeholders — those are filled by slide content
      const nvSpPr = g1(child, 'nvSpPr');
      const nvPr   = nvSpPr ? g1(nvSpPr, 'nvPr') : null;
      if (nvPr && g1(nvPr, 'ph')) continue;

      try {
        if      (ln === 'sp')           await renderShape(ctx, child, rels, images, this.themeColors, this.themeData, scale);
        else if (ln === 'pic')          await renderPicture(ctx, child, rels, images, this.themeColors, scale);
        else if (ln === 'grpSp')        await renderGroupShape(ctx, child, rels, images, this.themeColors, this.themeData, scale);
        else if (ln === 'graphicFrame') await renderGraphicFrame(ctx, child, this.themeColors, this.themeData, scale);
        else if (ln === 'cxnSp')        await renderConnector(ctx, child, this.themeColors, scale);
      } catch (e) {
        console.warn(`[PptxRenderer] master/layout shape error (${ln}):`, e);
      }
    }
  }

  // ── Metadata ─────────────────────────────────────────────────────────────

  /**
   * Parse animation steps for a slide.
   * Returns AnimStep[] sorted by clickNum then delay.
   * @param {number} slideIndex
   */
  getAnimations(slideIndex) {
    const slidePath = this.slidePaths[slideIndex];
    if (!slidePath || !this._files) return [];
    const raw = this._files[slidePath];
    if (!raw) return [];
    const slideDoc = parseXml(new TextDecoder().decode(raw));
    return parseAnimations(slideDoc);
  }

  /**
   * Parse slide transition info.
   * @param {number} slideIndex
   * @returns {{ type, duration, dir }|null}
   */
  getTransition(slideIndex) {
    const slidePath = this.slidePaths[slideIndex];
    if (!slidePath || !this._files) return null;
    const raw = this._files[slidePath];
    if (!raw) return null;
    const slideDoc = parseXml(new TextDecoder().decode(raw));
    return parseTransition(slideDoc);
  }

  /**
   * Create a PptxPlayer bound to this renderer and a canvas element.
   * Drives animation playback and slide transitions.
   *
   * @param {HTMLCanvasElement} canvas
   * @returns {PptxPlayer}
   * @example
   * const player = renderer.createPlayer(canvas);
   * await player.loadSlide(0);
   * player.play();
   * nextBtn.onclick = () => player.nextClick();
   */
  createPlayer(canvas) {
    return new PptxPlayer(this, canvas);
  }

  // ── SVG Export ──────────────────────────────────────────────────────────────

  /**
   * Render a slide to an SVG string.
   * SVG output has vector text (searchable), inline base64 images,
   * and proper gradient fills — matches PowerPoint's "Save as SVG".
   *
   * @param {number} slideIndex
   * @returns {Promise<string>}  complete SVG markup
   */
  async toSvg(slideIndex) {
    return renderSlideToSvg(slideIndex, this);
  }

  /**
   * Render all slides to SVG strings.
   * @returns {Promise<string[]>}
   */
  async allToSvg() {
    return renderAllSlidesToSvg(this);
  }

  // ── Embedded fonts ──────────────────────────────────────────────────────────

  /**
   * Decode and load any embedded fonts from the PPTX.
   * Fonts are loaded via FontFace API and become available to the renderer.
   *
   * @returns {Promise<EmbeddedFontResult[]>}  per-variant load results
   */
  async loadEmbeddedFonts() {
    const presRels = this._presRels || {};
    return loadEmbeddedFonts(this._files, presRels);
  }

  /**
   * List embedded fonts without loading them.
   * @returns {EmbeddedFontInfo[]}
   */
  listEmbeddedFonts() {
    return listEmbeddedFonts(this._files);
  }

  // ── Text extraction ─────────────────────────────────────────────────────────

  /**
   * Extract structured content from a slide.
   * Returns title, body text, tables, chart series names, alt text.
   * @param {number} slideIndex
   * @returns {Promise<SlideContent>}
   */
  async extractSlide(slideIndex) {
    return extractSlide(slideIndex, this);
  }

  /**
   * Extract content from all slides.
   * @returns {Promise<SlideContent[]>}
   */
  async extractAll() {
    return extractAll(this);
  }

  /**
   * Get all text from a slide as a plain string.
   * @param {number} slideIndex
   * @returns {Promise<string>}
   */
  async extractText(slideIndex) {
    return extractText(slideIndex, this);
  }

  /**
   * Full-text search across all slides.
   * @param {string} query
   * @returns {Promise<SearchResult[]>}
   */
  async searchSlides(query) {
    return searchSlides(query, this);
  }

  // ── Slide show ──────────────────────────────────────────────────────────────

  /**
   * Create a full-screen slide show player.
   *
   * @param {HTMLElement} container  — DOM element to mount into
   * @param {object}      [opts]     — SlideShow options
   * @returns {SlideShow}
   *
   * @example
   * const show = renderer.createShow(document.body, { showNotes: true });
   * await show.start(0);
   * // keyboard: ←→ Space PageUp/Down Home End Esc F
   */
  createShow(container, opts = {}) {
    return new SlideShow(this, container, opts);
  }

  // ── Clipboard / download ────────────────────────────────────────────────────

  /**
   * Copy a slide as a PNG image to the system clipboard.
   * @param {number} slideIndex
   * @param {object} [opts]
   * @param {number} [opts.dpi=150]   dots per inch
   * @param {number} [opts.width]    pixel width override
   * @returns {Promise<{success, method, dataUrl}>}
   */
  async copySlide(slideIndex, opts = {}) {
    return copySlideToClipboard(slideIndex, this, opts);
  }

  /**
   * Download a slide as a PNG file.
   * @param {number} slideIndex
   * @param {object} [opts]
   * @param {number} [opts.dpi=300]    dots per inch
   * @param {number} [opts.width]     pixel width override
   * @param {string} [opts.filename]
   */
  async downloadSlide(slideIndex, opts = {}) {
    return downloadSlide(slideIndex, this, opts);
  }

  /**
   * Download all slides as PNG files.
   * @param {object}   [opts]
   * @param {number}   [opts.dpi=300]      dots per inch
   * @param {number}   [opts.width]        pixel width override
   * @param {function} [opts.onProgress]   (completed, total) => void
   */
  async downloadAllSlides(opts = {}) {
    return downloadAllSlides(this, opts);
  }

  // ── Progressive deck view ───────────────────────────────────────────────────

  /**
   * Create a scrollable deck view with progressive lazy rendering.
   * Slides render on-demand as they scroll into the viewport.
   *
   * @param {HTMLElement} container
   * @param {object}      [opts]     — LazyDeckOpts
   * @returns {LazyDeckController}
   *
   * @example
   * const deck = renderer.createDeck(document.getElementById('viewer'));
   * // Scroll to slide 5:
   * deck.scrollTo(5);
   * // Force render everything:
   * await deck.renderAll();
   * // Clean up:
   * deck.destroy();
   */
  createDeck(container, opts = {}) {
    return createLazyDeck(this, container, opts);
  }

  // ── PPTX editing ────────────────────────────────────────────────────────────

  /**
   * Create a PptxWriter from this renderer for editing and re-export.
   *
   * @returns {PptxWriter}
   *
   * @example
   * const writer = renderer.edit();
   * writer.applyTemplate({ company: 'Acme', year: '2025' });
   * writer.setShapeText(0, 'Title 1', 'New Title');
   * writer.duplicateSlide(0);
   * await writer.download('edited.pptx');
   */
  edit() {
    return PptxWriter.fromRenderer(this);
  }

  // ── PDF export ──────────────────────────────────────────────────────────────

  /**
   * Export all slides (or a subset) to a PDF binary.
   *
   * @param {object}   [opts]
   * @param {number}   [opts.dpi=150]       dots per inch (96=screen, 150=default, 300=print)
   * @param {number}   [opts.width]         pixel width — overrides dpi if set
   * @param {number}   [opts.quality=0.92]  JPEG quality 0..1
   * @param {number[]} [opts.slides]        slide indices (default: all)
   * @param {function} [opts.onProgress]    (done, total) => void
   * @returns {Promise<Uint8Array>}
   *
   * @example
   * const bytes = await renderer.toPdf({ width: 2560, quality: 0.95 });
   * const blob = new Blob([bytes], { type: 'application/pdf' });
   */
  async toPdf(opts = {}) {
    return exportToPdf(this, opts);
  }

  /**
   * Export and download as a PDF file.
   * @param {string} [filename='presentation.pdf']
   * @param {object} [opts]
   */
  async downloadPdf(filename = 'presentation.pdf', opts = {}) {
    return downloadAsPdf(this, filename, opts);
  }

  /**
   * Export a single slide to PDF bytes.
   * @param {number} slideIndex
   * @param {object} [opts]
   * @returns {Promise<Uint8Array>}
   */
  async slideToPdf(slideIndex, opts = {}) {
    return exportSlideToPdf(slideIndex, this, opts);
  }

  /**
   * Get the speaker notes for a slide as plain text.
   * @param {number} slideIndex
   * @returns {Promise<string>} speaker notes, or empty string
   */
  async getSlideNotes(slideIndex) {
    if (slideIndex < 0 || slideIndex >= this.slidePaths.length) return '';
    const slideRels = await getRels(this._files, this.slidePaths[slideIndex]);
    const notesRel = Object.values(slideRels).find(r => r.type?.includes('notesSlide'));
    if (!notesRel) return '';
    const notesXml = this._readText(notesRel.fullPath);
    if (!notesXml) return '';
    const doc = parseXml(notesXml);
    // Collect all <a:t> text runs in notes, excluding the slide number placeholder
    const texts = [];
    for (const sp of gtn(doc, 'sp')) {
      const nvPr = g1(g1(sp, 'nvSpPr'), 'nvPr');
      const ph = nvPr ? g1(nvPr, 'ph') : null;
      // Skip slide number placeholder (type=sldNum)
      if (ph && (attr(ph, 'type') === 'sldNum')) continue;
      for (const t of gtn(sp, 't')) {
        texts.push(t.textContent);
      }
    }
    return texts.join('').trim();
  }

  /**
   * Get basic metadata about the presentation.
   * @returns {{ slideCount: number, width: number, height: number, widthEmu: number, heightEmu: number }}
   */
  getInfo() {
    return {
      slideCount: this.slideCount,
      widthEmu:  this.slideSize.cx,
      heightEmu: this.slideSize.cy,
      /** Slide width in inches (1 EMU = 1/914400 inch) */
      width:  this.slideSize.cx / 914400,
      /** Slide height in inches */
      height: this.slideSize.cy / 914400,
      /** Aspect ratio (width / height) */
      aspectRatio: this.slideSize.cx / this.slideSize.cy,
    };
  }

  /**
   * Render a slide and return a data URL (PNG by default).
   * @param {number} slideIndex
   * @param {number} [width=1280]
   * @param {string} [format='image/png']
   * @param {number} [quality=0.92]  used for image/jpeg
   * @returns {Promise<string>} data URL
   */
  async toDataURL(slideIndex, width = 1280, format = 'image/png', quality = 0.92) {
    const canvas = typeof OffscreenCanvas !== 'undefined'
      ? new OffscreenCanvas(1, 1)
      : document.createElement('canvas');
    await this.renderSlide(slideIndex, canvas, width);
    if (canvas instanceof OffscreenCanvas) {
      const blob = await canvas.convertToBlob({ type: format, quality });
      return new Promise(resolve => {
        const fr = new FileReader();
        fr.onload = () => resolve(fr.result);
        fr.readAsDataURL(blob);
      });
    }
    return canvas.toDataURL(format, quality);
  }

  /**
   * Render a slide and return a Blob.
   * @param {number} slideIndex
   * @param {number} [width=1280]
   * @param {string} [format='image/png']
   * @returns {Promise<Blob>}
   */
  async toBlob(slideIndex, width = 1280, format = 'image/png') {
    const canvas = typeof OffscreenCanvas !== 'undefined'
      ? new OffscreenCanvas(1, 1)
      : document.createElement('canvas');
    await this.renderSlide(slideIndex, canvas, width);
    if (canvas instanceof OffscreenCanvas) return canvas.convertToBlob({ type: format });
    return new Promise(resolve => canvas.toBlob(resolve, format));
  }

  // ── Lifecycle ─────────────────────────────────────────────────────────────

  /** Release all blob: URLs created during rendering. */
  destroy() {
    for (const url of this._blobUrls) {
      try { URL.revokeObjectURL(url); } catch (_) {}
    }
    this._blobUrls    = [];
    this._files       = {};
    this.masterDoc    = null;
    this.masterImages = {};
  }
}

export { PptxRenderer };

// Static font utilities — usable without a renderer instance
export {
  registerFont,
  registerFonts,
  listRegisteredFonts,
  clearRegisteredFonts,
  isFontAvailable,
} from './fonts.js';

// Animation / playback utilities
export {
  PptxPlayer,
  parseAnimations,
  parseTransition,
  renderTransitionFrame,
  compositeShape,
} from './animation.js';

// SVG export
export { renderSlideToSvg, renderAllSlidesToSvg } from './svg.js';

// Embedded font decoding
export { loadEmbeddedFonts, listEmbeddedFonts } from './fntdata.js';

// Text extraction / search
export { extractSlide, extractAll, extractText, searchSlides } from './extract.js';

// Slide show
export { SlideShow } from './slideshow.js';

// Clipboard / download / progressive rendering
export { copySlideToClipboard, downloadSlide, downloadAllSlides, createLazyDeck } from './clipboard.js';

// PPTX writer / editor
export { PptxWriter } from './writer.js';

// PDF export
export { exportToPdf, downloadAsPdf, exportSlideToPdf } from './pdf.js';
