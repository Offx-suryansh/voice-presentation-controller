import PptxRenderer from '../../vendor/pptx-browser/src/index.js';

/** Renders PDF and PPTX slides to the shared presentation canvas. */
export class PresentationViewer {
  constructor(canvas, stage) {
    Object.assign(this, { canvas, stage, doc: null, pptx: null, format: '', busy: false, queued: null });
  }
  get ready() { return !!(this.doc || this.pptx); }
  async load(file) {
    const invalid = () => Object.assign(new Error('invalid'), { code: 'invalid' });
    if (!file || !file.size) throw invalid();
    const extension = file.name.toLowerCase().split('.').pop();
    if (!['pdf', 'pptx'].includes(extension)) throw invalid();
    const buf = await file.arrayBuffer();
    const signature = new Uint8Array(buf.slice(0, 5));

    if (extension === 'pdf') {
      if (String.fromCharCode(...signature) !== '%PDF-') throw invalid();
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.js';
      const doc = await pdfjsLib.getDocument({ data: buf }).promise;
      this.close();
      this.doc = doc;
      this.format = 'pdf';
      return doc.numPages;
    }

    if (signature[0] !== 0x50 || signature[1] !== 0x4b) throw invalid();
    const pptx = new PptxRenderer();
    try {
      await pptx.load(buf);
      if (!pptx.slideCount) throw invalid();
      this.close();
      this.pptx = pptx;
      this.format = 'pptx';
      return pptx.slideCount;
    } catch (error) {
      pptx.destroy();
      throw error;
    }
  }
  close() {
    if (this.doc) this.doc.destroy();
    if (this.pptx) this.pptx.destroy();
    this.doc = null; this.pptx = null; this.format = ''; this.queued = null;
    this.canvas.width = 0; this.canvas.height = 0;
    this.canvas.style.width = ''; this.canvas.style.height = '';
  }
  async render(pageNo, zoom) {
    if (this.busy) { this.queued = [pageNo, zoom]; return; }
    this.busy = true;
    const source = this.doc || this.pptx;
    try {
      await new Promise(requestAnimationFrame); // let layout settle (e.g. after fullscreen)
      const doc = this.doc, pptx = this.pptx, c = this.canvas;
      if (doc) {
        const page = await doc.getPage(pageNo), base = page.getViewport({ scale: 1 });
        const fit = Math.min((this.stage.clientWidth - 32) / base.width, (this.stage.clientHeight - 32) / base.height);
        const dpr = window.devicePixelRatio || 1, vp = page.getViewport({ scale: Math.max(fit, 0.05) * zoom * dpr });
        c.width = vp.width; c.height = vp.height;
        c.style.width = vp.width / dpr + 'px'; c.style.height = vp.height / dpr + 'px';
        await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      } else if (pptx) {
        const ratio = pptx.slideSize.cy / pptx.slideSize.cx;
        const fitWidth = Math.min(this.stage.clientWidth - 32, (this.stage.clientHeight - 32) / ratio);
        const dpr = window.devicePixelRatio || 1;
        const renderWidth = Math.max(1, Math.round(fitWidth * zoom * dpr));
        await pptx.renderSlide(pageNo - 1, c, renderWidth);
        c.style.width = c.width / dpr + 'px'; c.style.height = c.height / dpr + 'px';
      }
    } catch (error) {
      if (source && (this.doc === source || this.pptx === source)) throw error;
    }
    finally {
      this.busy = false;
      if (this.queued) { const q = this.queued; this.queued = null; return this.render(...q); }
    }
  }
}
