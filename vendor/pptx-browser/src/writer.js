/**
 * writer.js — PPTX editor and template engine.
 *
 * Loads an existing PPTX (via a PptxRenderer instance or raw bytes),
 * provides a fluent API to mutate its content, and serializes back to
 * a valid PPTX file that PowerPoint / Keynote / LibreOffice can open.
 *
 * ── Quick start ──────────────────────────────────────────────────────────────
 *
 *   import { PptxWriter } from 'pptx-canvas-renderer';
 *
 *   // From a loaded renderer:
 *   const writer = PptxWriter.fromRenderer(renderer);
 *
 *   // Or from raw bytes:
 *   const writer = await PptxWriter.fromBytes(arrayBuffer);
 *
 *   // Template substitution  ({{tokens}} in shapes / speaker notes)
 *   writer.applyTemplate({ name: 'Acme Corp', year: '2025' });
 *
 *   // Replace text everywhere
 *   writer.replaceText('Old Text', 'New Text');
 *
 *   // Set the text of a specific shape
 *   writer.setShapeText(0, 'Title 1', 'My New Title');
 *
 *   // Swap an image on slide 2 (shape named "Picture 1")
 *   await writer.setShapeImage(1, 'Picture 1', jpegBytes, 'image/jpeg');
 *
 *   // Duplicate slide 0 as a new slide at the end
 *   writer.duplicateSlide(0);
 *
 *   // Remove slide 3
 *   writer.removeSlide(3);
 *
 *   // Reorder slides
 *   writer.reorderSlides([2, 0, 1]);
 *
 *   // Change a theme colour
 *   writer.setThemeColor('accent1', 'FF0000');
 *
 *   // Export
 *   const bytes = await writer.save();        // → Uint8Array
 *   writer.download('edited.pptx');           // trigger browser download
 *
 * ── API reference ─────────────────────────────────────────────────────────────
 *
 *   PptxWriter.fromRenderer(renderer)     — clone from loaded PptxRenderer
 *   PptxWriter.fromBytes(buffer)          — parse PPTX bytes fresh
 *
 *   .applyTemplate(data, opts)            — {{token}} substitution
 *   .replaceText(find, replace, opts)     — global find-and-replace
 *   .setShapeText(slideIdx, name, text)   — set text of named shape
 *   .getShapeText(slideIdx, name)         — read text from named shape
 *   .addTextBox(slideIdx, text, style)    — add a new text box
 *   .setShapeImage(slideIdx, name, bytes, mime)  — swap shape image
 *   .addImage(slideIdx, bytes, mime, rect)       — add new image shape
 *   .setSlideBackground(slideIdx, color)         — solid background color
 *   .setThemeColor(key, hexRgb)           — change theme colour (no #)
 *   .duplicateSlide(fromIdx, toIdx?)      — copy slide
 *   .removeSlide(slideIdx)                — delete slide
 *   .reorderSlides(newOrder)              — reorder by index array
 *   .setSlideNotes(slideIdx, text)        — set speaker notes
 *   .getSlidePaths()                      — current slide file paths
 *   .getSlideCount()                      — current slide count
 *   .save()                               → Promise<Uint8Array>  PPTX bytes
 *   .download(filename)                   — save file in browser
 */

import { readZip } from './zip.js';
import { writeZip } from './zip-writer.js';

const dec = new TextDecoder();
const enc = new TextEncoder();
const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};

// ── XML helpers ───────────────────────────────────────────────────────────────

function parseXml(str) {
  return new DOMParser().parseFromString(str, 'application/xml');
}
function serializeXml(doc) {
  const s = new XMLSerializer().serializeToString(doc);
  // Ensure declaration
  if (s.startsWith('<?xml')) return s;
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
}
function xmlBytes(doc) { return enc.encode(serializeXml(doc)); }
function readXml(files, path) {
  const raw = files[path];
  if (!raw) return null;
  return parseXml(dec.decode(raw));
}
function g1(node, name) {
  if (!node) return null;
  const all = node.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) return all[i];
  return null;
}
function gtn(node, name) {
  if (!node) return [];
  const r = [];
  const all = node.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) r.push(all[i]);
  return r;
}
function attr(el, name, def = null) {
  if (!el) return def;
  const v = el.getAttribute(name);
  return v !== null ? v : def;
}

// ── Relationship helpers ──────────────────────────────────────────────────────

function relsPath(filePath) {
  const parts = filePath.split('/');
  const name  = parts.pop();
  return [...parts, '_rels', name + '.rels'].join('/');
}

function parseRels(files, filePath) {
  const doc = readXml(files, relsPath(filePath));
  if (!doc) return {};
  const map = {};
  for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id     = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    const type   = rel.getAttribute('Type') || '';
    let fullPath = target;
    if (!target.startsWith('/') && !target.startsWith('http')) {
      const dir = filePath.split('/').slice(0, -1).join('/');
      fullPath  = dir ? dir + '/' + target.replace(/^\.\.\//, '') : target;
      // Handle ../ traversal
      const parts = fullPath.split('/');
      const resolved = [];
      for (const p of parts) {
        if (p === '..') resolved.pop();
        else resolved.push(p);
      }
      fullPath = resolved.join('/');
    }
    map[id] = { id, target, type, fullPath };
  }
  return map;
}

function buildRelsDoc(rels) {
  const doc = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
  const root = doc.documentElement;
  for (const rel of Object.values(rels)) {
    const el = doc.createElementNS(NS.rel, 'Relationship');
    el.setAttribute('Id', rel.id);
    el.setAttribute('Type', rel.type);
    el.setAttribute('Target', rel.target);
    if (rel.targetMode) el.setAttribute('TargetMode', rel.targetMode);
    root.appendChild(el);
  }
  return doc;
}

function nextRId(rels) {
  const nums = Object.keys(rels)
    .map(id => parseInt(id.replace('rId', ''), 10))
    .filter(n => !isNaN(n));
  return 'rId' + ((nums.length ? Math.max(...nums) : 0) + 1);
}

// ── Shape lookup helpers ──────────────────────────────────────────────────────

function findShapeByName(spTree, name) {
  for (const child of spTree.children) {
    const ln = child.localName;
    if (ln === 'sp' || ln === 'pic' || ln === 'cxnSp') {
      const nvEl  = g1(child, 'nvSpPr') || g1(child, 'nvPicPr') || g1(child, 'nvCxnSpPr');
      const cNvPr = nvEl ? g1(nvEl, 'cNvPr') : null;
      if (cNvPr) {
        const shapeName = cNvPr.getAttribute('name') || '';
        if (shapeName === name) return child;
      }
    } else if (ln === 'grpSp') {
      const found = findShapeByName(child, name);
      if (found) return found;
    }
  }
  return null;
}

function findShapeById(spTree, id) {
  const idStr = String(id);
  for (const child of spTree.children) {
    const ln = child.localName;
    if (ln === 'sp' || ln === 'pic' || ln === 'cxnSp') {
      const nvEl  = g1(child, 'nvSpPr') || g1(child, 'nvPicPr');
      const cNvPr = nvEl ? g1(nvEl, 'cNvPr') : null;
      if (cNvPr && (cNvPr.getAttribute('id') || '') === idStr) return child;
    }
  }
  return null;
}

function getSpTree(slideDoc) {
  const cSld = g1(slideDoc, 'cSld');
  return cSld ? g1(cSld, 'spTree') : null;
}

// ── Text replacement helpers ──────────────────────────────────────────────────

function getAllTextNodes(node) {
  const result = [];
  const walker = node.ownerDocument
    ? node.ownerDocument.createTreeWalker(node, 0x04 /* NodeFilter.SHOW_TEXT */)
    : null;
  if (!walker) return result;
  let n;
  while ((n = walker.nextNode())) result.push(n);
  return result;
}

/** Replace text in a run without disturbing formatting. */
function replaceInDoc(doc, find, replace, caseSensitive = true) {
  // Collect all <a:t> elements and replace within their text content
  for (const t of gtn(doc, 't')) {
    const orig = t.textContent;
    if (!orig) continue;
    const newText = caseSensitive
      ? orig.split(find).join(replace)
      : orig.replace(new RegExp(escapeRegex(find), 'gi'), replace);
    if (newText !== orig) t.textContent = newText;
  }
}

function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

// ── Content type helpers ──────────────────────────────────────────────────────

const MIME_EXT = {
  'image/jpeg': 'jpeg', 'image/jpg': 'jpeg',
  'image/png': 'png', 'image/gif': 'gif',
  'image/webp': 'webp', 'image/svg+xml': 'svg',
};
const CT_MAP = {
  jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif',
  webp: 'image/webp', svg: 'image/svg+xml',
};

function addContentType(files, ext, partName) {
  const ctPath = '[Content_Types].xml';
  const doc    = readXml(files, ctPath);
  if (!doc) return;
  const root = doc.documentElement;
  // Check if Override already exists
  for (const ov of gtn(doc, 'Override')) {
    if (ov.getAttribute('PartName') === '/' + partName) return;
  }
  const ov = doc.createElementNS(NS.ct, 'Override');
  ov.setAttribute('PartName',    '/' + partName);
  ov.setAttribute('ContentType', CT_MAP[ext] || 'application/octet-stream');
  root.appendChild(ov);
  files[ctPath] = xmlBytes(doc);
}

// ── PptxWriter ────────────────────────────────────────────────────────────────

export class PptxWriter {
  constructor(files) {
    /** @private Mutable copy of all ZIP entries */
    this._files = files;

    // Parse presentation.xml once
    this._presPath = 'ppt/presentation.xml';
    this._presDoc  = readXml(files, this._presPath);
    this._presRels = parseRels(files, this._presPath);

    // Build ordered slide path list
    this._slidePaths = this._buildSlidePaths();
  }

  // ── Factory ─────────────────────────────────────────────────────────────────

  /** Clone from an already-loaded PptxRenderer. O(1) — shares byte arrays. */
  static fromRenderer(renderer) {
    // Deep-copy the files map so mutations don't affect the renderer
    const files = {};
    for (const [k, v] of Object.entries(renderer._files)) {
      files[k] = v instanceof Uint8Array ? v.slice() : v;
    }
    return new PptxWriter(files);
  }

  /** Parse from raw ArrayBuffer or Uint8Array. */
  static async fromBytes(buffer) {
    const files = await readZip(buffer);
    return new PptxWriter(files);
  }

  // ── Slide list ──────────────────────────────────────────────────────────────

  _buildSlidePaths() {
    if (!this._presDoc) return [];
    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (!sldIdLst) return [];
    const paths = [];
    for (const sldId of sldIdLst.children) {
      if (sldId.localName !== 'sldId') continue;
      const rId = sldId.getAttribute('r:id') || sldId.getAttribute('id');
      const rel = this._presRels[rId];
      if (rel) paths.push(rel.fullPath);
    }
    return paths;
  }

  _savePresDoc() {
    this._files[this._presPath] = xmlBytes(this._presDoc);
  }

  _savePresRels() {
    this._files[relsPath(this._presPath)] = xmlBytes(buildRelsDoc(this._presRels));
  }

  getSlidePaths()  { return [...this._slidePaths]; }
  getSlideCount()  { return this._slidePaths.length; }

  _slideDoc(idx) {
    const path = this._slidePaths[idx];
    if (!path) throw new RangeError(`Slide ${idx} out of range`);
    return readXml(this._files, path);
  }

  _saveSlideDoc(idx, doc) {
    this._files[this._slidePaths[idx]] = xmlBytes(doc);
  }

  // ── Template substitution ────────────────────────────────────────────────────

  /**
   * Replace `{{key}}` placeholders with values from a data object.
   * Applied to every slide, every text shape, and speaker notes.
   *
   * @param {Record<string, string|number>} data
   * @param {object} [opts]
   * @param {string} [opts.open='{{']
   * @param {string} [opts.close='}}']
   * @param {number[]} [opts.slides]  limit to specific slide indices
   */
  applyTemplate(data, opts = {}) {
    const { open = '{{', close = '}}', slides } = opts;
    const indices = slides ?? this._slidePaths.map((_, i) => i);

    for (const idx of indices) {
      const doc = this._slideDoc(idx);

      for (const [key, value] of Object.entries(data)) {
        const token = open + key + close;
        replaceInDoc(doc, token, String(value));
      }

      this._saveSlideDoc(idx, doc);
    }

    // Also apply to speaker notes
    for (const idx of indices) {
      this._applyTemplateToNotes(idx, data, open, close);
    }
  }

  _applyTemplateToNotes(idx, data, open, close) {
    const slideRels = parseRels(this._files, this._slidePaths[idx]);
    const notesRel  = Object.values(slideRels).find(r => r.type?.includes('notesSlide'));
    if (!notesRel) return;
    const notesDoc = readXml(this._files, notesRel.fullPath);
    if (!notesDoc) return;
    for (const [key, value] of Object.entries(data)) {
      replaceInDoc(notesDoc, open + key + close, String(value));
    }
    this._files[notesRel.fullPath] = xmlBytes(notesDoc);
  }

  // ── Global find-and-replace ──────────────────────────────────────────────────

  /**
   * Find and replace text across all (or specified) slides.
   * @param {string} find
   * @param {string} replace
   * @param {object} [opts]
   * @param {boolean} [opts.caseSensitive=true]
   * @param {boolean} [opts.includeNotes=false]
   * @param {number[]} [opts.slides]
   */
  replaceText(find, replace, opts = {}) {
    const { caseSensitive = true, includeNotes = false, slides } = opts;
    const indices = slides ?? this._slidePaths.map((_, i) => i);

    for (const idx of indices) {
      const doc = this._slideDoc(idx);
      replaceInDoc(doc, find, replace, caseSensitive);
      this._saveSlideDoc(idx, doc);

      if (includeNotes) {
        const slideRels = parseRels(this._files, this._slidePaths[idx]);
        const notesRel  = Object.values(slideRels).find(r => r.type?.includes('notesSlide'));
        if (notesRel) {
          const nd = readXml(this._files, notesRel.fullPath);
          if (nd) {
            replaceInDoc(nd, find, replace, caseSensitive);
            this._files[notesRel.fullPath] = xmlBytes(nd);
          }
        }
      }
    }
  }

  // ── Shape text ───────────────────────────────────────────────────────────────

  /**
   * Set the text content of a named shape on a slide.
   * Preserves the formatting of the first run; clears all other runs.
   *
   * @param {number} slideIdx
   * @param {string} shapeName     exact `name` attribute of the shape
   * @param {string} text          new text (use \n for line breaks)
   * @param {object} [opts]
   * @param {boolean} [opts.preserveFormatting=true]
   */
  setShapeText(slideIdx, shapeName, text, opts = {}) {
    const { preserveFormatting = true } = opts;
    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const shape = findShapeByName(spTree, shapeName);
    if (!shape) throw new Error(`Shape "${shapeName}" not found on slide ${slideIdx}`);

    const txBody = g1(shape, 'txBody');
    if (!txBody) return this;

    // Get reference run properties
    const firstRun = g1(txBody, 'r');
    const refRPr   = firstRun ? g1(firstRun, 'rPr') : null;
    const refPPr   = g1(g1(txBody, 'p'), 'pPr');

    // Remove all existing paragraphs
    for (const p of gtn(txBody, 'p')) p.parentNode.removeChild(p);

    const lines = text.split('\n');
    const nsA   = NS.a;

    for (const line of lines) {
      const p = doc.createElementNS(nsA, 'a:p');

      if (refPPr && preserveFormatting) {
        p.appendChild(refPPr.cloneNode(true));
      }

      const r  = doc.createElementNS(nsA, 'a:r');
      if (refRPr && preserveFormatting) {
        r.appendChild(refRPr.cloneNode(true));
      }
      const t = doc.createElementNS(nsA, 'a:t');
      t.textContent = line;
      r.appendChild(t);
      p.appendChild(r);
      txBody.appendChild(p);
    }

    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  /**
   * Read the plain text of a named shape.
   * @param {number} slideIdx
   * @param {string} shapeName
   * @returns {string}
   */
  getShapeText(slideIdx, shapeName) {
    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return '';
    const shape  = findShapeByName(spTree, shapeName);
    if (!shape) return '';
    return gtn(shape, 't').map(t => t.textContent).join('');
  }

  // ── Add text box ─────────────────────────────────────────────────────────────

  /**
   * Add a new text box to a slide.
   *
   * @param {number} slideIdx
   * @param {string} text
   * @param {object} style
   * @param {number} style.x      EMU from left edge
   * @param {number} style.y      EMU from top edge
   * @param {number} style.w      EMU width
   * @param {number} style.h      EMU height
   * @param {string} [style.color]      hex colour, no #
   * @param {number} [style.fontSize]   pt * 100  (e.g. 2400 = 24pt)
   * @param {boolean}[style.bold]
   * @param {string} [style.align]      l|ctr|r
   * @param {string} [style.fontFamily]
   */
  addTextBox(slideIdx, text, style = {}) {
    const {
      x = 914400, y = 914400, w = 4572000, h = 914400,
      color = '000000', fontSize = 1800, bold = false,
      align = 'l', fontFamily = 'Calibri',
    } = style;

    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    // Next shape ID
    const maxId = Math.max(0, ...gtn(spTree, 'cNvPr').map(e => parseInt(e.getAttribute('id') || '0', 10)));
    const newId = maxId + 1;
    const name  = `TextBox ${newId}`;

    const nsA = NS.a, nsP = NS.p;

    const xml = `<p:sp xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:nvSpPr>
    <p:cNvPr id="${newId}" name="${name}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="${align}"/>
      <a:r>
        <a:rPr lang="en-US" sz="${fontSize}" b="${bold ? 1 : 0}" dirty="0">
          <a:solidFill><a:srgbClr val="${color}"/></a:solidFill>
          <a:latin typeface="${fontFamily}"/>
        </a:rPr>
        <a:t>${text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

    const frag = parseXml(xml);
    spTree.appendChild(doc.adoptNode(frag.documentElement));
    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  // ── Image replacement ─────────────────────────────────────────────────────────

  /**
   * Replace the image in a named picture shape.
   *
   * @param {number}     slideIdx
   * @param {string}     shapeName
   * @param {Uint8Array} imageBytes
   * @param {string}     [mimeType='image/jpeg']
   */
  async setShapeImage(slideIdx, shapeName, imageBytes, mimeType = 'image/jpeg') {
    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const shape = findShapeByName(spTree, shapeName);
    if (!shape) throw new Error(`Shape "${shapeName}" not found on slide ${slideIdx}`);

    const slideRels = parseRels(this._files, this._slidePaths[slideIdx]);

    // Find existing blip rId
    const blipFill = g1(shape, 'blipFill');
    const blip     = blipFill ? g1(blipFill, 'blip') : null;
    const oldRId   = blip ? (blip.getAttribute('r:embed') || blip.getAttribute('embed')) : null;
    const oldRel   = oldRId ? slideRels[oldRId] : null;

    // Write new media file
    const ext      = MIME_EXT[mimeType] || 'jpeg';
    const mediaIdx = Object.keys(this._files).filter(p => p.startsWith('ppt/media/')).length + 1;
    const mediaPath = `ppt/media/image${mediaIdx}.${ext}`;
    this._files[mediaPath] = imageBytes;

    // Update or create relationship
    let rId;
    if (oldRId && oldRel) {
      // Reuse the old rId, just point to the new file
      rId = oldRId;
      slideRels[rId] = {
        id: rId, type: oldRel.type,
        target: `../media/image${mediaIdx}.${ext}`,
        fullPath: mediaPath,
      };
    } else {
      rId = nextRId(slideRels);
      slideRels[rId] = {
        id: rId,
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        target: `../media/image${mediaIdx}.${ext}`,
        fullPath: mediaPath,
      };
    }

    // Update the blip element
    if (blip) {
      blip.setAttribute('r:embed', rId);
    }

    // Save updated rels and slide doc
    this._files[relsPath(this._slidePaths[slideIdx])] = xmlBytes(buildRelsDoc(slideRels));
    this._saveSlideDoc(slideIdx, doc);

    // Add content type
    addContentType(this._files, ext, mediaPath);
    return this;
  }

  /**
   * Add a new image shape to a slide.
   *
   * @param {number}     slideIdx
   * @param {Uint8Array} imageBytes
   * @param {string}     [mimeType='image/jpeg']
   * @param {object}     rect      { x, y, w, h } in EMU
   */
  async addImage(slideIdx, imageBytes, mimeType = 'image/jpeg', rect = {}) {
    const {
      x = 914400, y = 914400,
      w = 2743200, h = 2057400,
    } = rect;

    const ext       = MIME_EXT[mimeType] || 'jpeg';
    const mediaIdx  = Object.keys(this._files).filter(p => p.startsWith('ppt/media/')).length + 1;
    const mediaPath = `ppt/media/image${mediaIdx}.${ext}`;
    this._files[mediaPath] = imageBytes;

    const slideRels = parseRels(this._files, this._slidePaths[slideIdx]);
    const rId = nextRId(slideRels);
    slideRels[rId] = {
      id: rId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: `../media/image${mediaIdx}.${ext}`,
      fullPath: mediaPath,
    };
    this._files[relsPath(this._slidePaths[slideIdx])] = xmlBytes(buildRelsDoc(slideRels));

    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const maxId = Math.max(0, ...gtn(spTree, 'cNvPr').map(e => parseInt(e.getAttribute('id') || '0', 10)));
    const newId = maxId + 1;
    const nsA = NS.a, nsP = NS.p;

    const xml = `<p:pic xmlns:p="${nsP}" xmlns:a="${nsA}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvPicPr>
    <p:cNvPr id="${newId}" name="Picture ${newId}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="${rId}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;

    const frag = parseXml(xml);
    spTree.appendChild(doc.adoptNode(frag.documentElement));
    this._saveSlideDoc(slideIdx, doc);
    addContentType(this._files, ext, mediaPath);
    return this;
  }

  // ── Slide background ─────────────────────────────────────────────────────────

  /**
   * Set a solid colour background on a slide.
   * @param {number} slideIdx
   * @param {string} hexRgb   6-digit hex, no '#'
   */
  setSlideBackground(slideIdx, hexRgb) {
    const doc  = this._slideDoc(slideIdx);
    const cSld = g1(doc, 'cSld');
    if (!cSld) return this;

    // Remove existing bg
    const oldBg = g1(cSld, 'bg');
    if (oldBg) cSld.removeChild(oldBg);

    const nsA = NS.a, nsP = NS.p;
    const xml = `<p:bg xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:bgPr><a:solidFill><a:srgbClr val="${hexRgb}"/></a:solidFill>
  <a:effectLst/></p:bgPr></p:bg>`;
    const bgEl = doc.adoptNode(parseXml(xml).documentElement);
    // Insert bg as first child of cSld
    cSld.insertBefore(bgEl, cSld.firstChild);
    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  // ── Theme colours ─────────────────────────────────────────────────────────────

  /**
   * Override a theme colour.
   * Key: dk1|lt1|dk2|lt2|accent1…accent6|hlink|folHlink
   * Value: 6-digit hex RGB, no '#'
   *
   * @param {string} key
   * @param {string} hexRgb
   */
  setThemeColor(key, hexRgb) {
    // Find theme file via presentation rels
    const presRels = this._presRels;
    let themePath  = Object.values(presRels).find(r => r.type?.includes('theme'))?.fullPath;

    if (!themePath) {
      // Try via slide master
      const masterRel = Object.values(presRels).find(r => r.type?.includes('slideMaster'));
      if (masterRel) {
        const mr = parseRels(this._files, masterRel.fullPath);
        themePath = Object.values(mr).find(r => r.type?.includes('theme'))?.fullPath;
      }
    }
    if (!themePath) return this;

    const doc = readXml(this._files, themePath);
    if (!doc) return this;

    // Map theme key to element path: e.g. accent1 → a:accent1 > a:srgbClr
    const fmtScheme = g1(doc, 'fmtScheme');
    const clrScheme = g1(doc, 'clrScheme');
    if (!clrScheme) return this;

    // Find the element with matching local name
    for (const child of clrScheme.children) {
      if (child.localName === key) {
        // Replace or set inner colour element
        const srgb = g1(child, 'srgbClr');
        if (srgb) {
          srgb.setAttribute('val', hexRgb);
        } else {
          while (child.firstChild) child.removeChild(child.firstChild);
          const nsA = NS.a;
          const el  = doc.createElementNS(nsA, 'a:srgbClr');
          el.setAttribute('val', hexRgb);
          child.appendChild(el);
        }
        break;
      }
    }

    this._files[themePath] = xmlBytes(doc);
    return this;
  }

  // ── Slide operations ──────────────────────────────────────────────────────────

  /**
   * Duplicate a slide.
   * @param {number} fromIdx       source slide index
   * @param {number} [toIdx]       insert position (default: end)
   */
  duplicateSlide(fromIdx, toIdx) {
    const insertAt = toIdx ?? this._slidePaths.length;
    const srcPath  = this._slidePaths[fromIdx];
    if (!srcPath) throw new RangeError(`Slide ${fromIdx} out of range`);

    // Find next available slide number
    const nums = Object.keys(this._files)
      .map(p => p.match(/ppt\/slides\/slide(\d+)\.xml/))
      .filter(Boolean).map(m => parseInt(m[1], 10));
    const nextNum = (nums.length ? Math.max(...nums) : 0) + 1;

    const newSlidePath = `ppt/slides/slide${nextNum}.xml`;
    const newRelsPath  = relsPath(newSlidePath);

    // Copy slide XML
    this._files[newSlidePath] = this._files[srcPath].slice();

    // Copy slide rels (images etc. are shared)
    const srcRelsPath = relsPath(srcPath);
    if (this._files[srcRelsPath]) {
      this._files[newRelsPath] = this._files[srcRelsPath].slice();
    }

    // Add relationship in presentation.xml.rels
    const newRId = nextRId(this._presRels);
    const target = `slides/slide${nextNum}.xml`;
    this._presRels[newRId] = {
      id: newRId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      target,
      fullPath: newSlidePath,
    };
    this._savePresRels();

    // Add sldId to presentation.xml sldIdLst
    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (sldIdLst) {
      const ids = gtn(sldIdLst, 'sldId').map(el => parseInt(el.getAttribute('id') || '0', 10));
      const nextId = (ids.length ? Math.max(...ids) : 255) + 1;
      const nsP  = NS.p;
      const sldIdEl = this._presDoc.createElementNS(nsP, 'p:sldId');
      sldIdEl.setAttribute('id', String(nextId));
      sldIdEl.setAttributeNS(NS.r, 'r:id', newRId);

      // Insert at correct position
      const children = Array.from(sldIdLst.children);
      if (insertAt >= children.length) {
        sldIdLst.appendChild(sldIdEl);
      } else {
        sldIdLst.insertBefore(sldIdEl, children[insertAt]);
      }
    }
    this._savePresDoc();

    // Rebuild slide path list
    this._slidePaths = this._buildSlidePaths();

    // Add content type for new slide
    const ctPath = '[Content_Types].xml';
    const ctDoc  = readXml(this._files, ctPath);
    if (ctDoc) {
      const root = ctDoc.documentElement;
      const ov   = ctDoc.createElementNS(NS.ct, 'Override');
      ov.setAttribute('PartName',    '/' + newSlidePath);
      ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
      root.appendChild(ov);
      this._files[ctPath] = xmlBytes(ctDoc);
    }

    return this;
  }

  /**
   * Remove a slide.
   * @param {number} slideIdx
   */
  removeSlide(slideIdx) {
    if (this._slidePaths.length <= 1) throw new Error('Cannot remove the last slide');
    const path = this._slidePaths[slideIdx];
    if (!path) throw new RangeError(`Slide ${slideIdx} out of range`);

    // Remove from sldIdLst
    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (sldIdLst) {
      for (const sldId of Array.from(sldIdLst.children)) {
        const rId = sldId.getAttribute('r:id') || sldId.getAttribute('id');
        const rel = this._presRels[rId];
        if (rel && rel.fullPath === path) {
          sldIdLst.removeChild(sldId);
          delete this._presRels[rId];
          break;
        }
      }
    }
    this._savePresDoc();
    this._savePresRels();
    this._slidePaths = this._buildSlidePaths();
    return this;
  }

  /**
   * Reorder slides.
   * @param {number[]} newOrder  e.g. [2, 0, 1] to put slide 2 first
   */
  reorderSlides(newOrder) {
    if (newOrder.length !== this._slidePaths.length) {
      throw new Error('newOrder must have the same length as the current slide count');
    }

    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (!sldIdLst) return this;

    const children = Array.from(sldIdLst.children).filter(el => el.localName === 'sldId');
    // Detach all
    for (const c of children) sldIdLst.removeChild(c);
    // Re-attach in new order
    for (const idx of newOrder) {
      if (children[idx]) sldIdLst.appendChild(children[idx]);
    }
    this._savePresDoc();
    this._slidePaths = this._buildSlidePaths();
    return this;
  }

  // ── Speaker notes ─────────────────────────────────────────────────────────────

  /**
   * Set the speaker notes for a slide. Creates the notes slide if absent.
   * @param {number} slideIdx
   * @param {string} text
   */
  setSlideNotes(slideIdx, text) {
    const slidePath = this._slidePaths[slideIdx];
    const slideRels = parseRels(this._files, slidePath);
    const notesRel  = Object.values(slideRels).find(r => r.type?.includes('notesSlide'));

    if (notesRel) {
      const nd = readXml(this._files, notesRel.fullPath);
      if (nd) {
        for (const sp of gtn(nd, 'sp')) {
          const nvPr = g1(g1(sp, 'nvSpPr'), 'nvPr');
          const ph   = nvPr ? g1(nvPr, 'ph') : null;
          if (ph && attr(ph, 'type') !== 'sldNum') {
            for (const t of gtn(sp, 't')) t.textContent = '';
            const firstT = g1(sp, 't');
            if (firstT) firstT.textContent = text;
            break;
          }
        }
        this._files[notesRel.fullPath] = xmlBytes(nd);
      }
    } else {
      // Create a minimal notes slide
      this._createNotesSlide(slideIdx, slidePath, slideRels, text);
    }
    return this;
  }

  _createNotesSlide(slideIdx, slidePath, slideRels, text) {
    const num = Object.keys(this._files).filter(p => p.startsWith('ppt/notesSlides/')).length + 1;
    const nsP = NS.p, nsA = NS.a;
    const notesPath = `ppt/notesSlides/notesSlide${num}.xml`;

    const notesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/>
        <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
        <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:t>${text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>`;

    this._files[notesPath] = enc.encode(notesXml);

    const newRId = nextRId(slideRels);
    slideRels[newRId] = {
      id: newRId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
      target: `../notesSlides/notesSlide${num}.xml`,
      fullPath: notesPath,
    };
    this._files[relsPath(slidePath)] = xmlBytes(buildRelsDoc(slideRels));
  }

  // ── Serialisation ─────────────────────────────────────────────────────────────

  /**
   * Serialize the edited PPTX to bytes.
   * @returns {Promise<Uint8Array>}
   */
  async save() {
    return writeZip(this._files);
  }

  /**
   * Download as a PPTX file in the browser.
   * @param {string} [filename='edited.pptx']
   */
  async download(filename = 'edited.pptx') {
    const bytes = await this.save();
    const blob  = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  // ── Utility ───────────────────────────────────────────────────────────────────

  /**
   * List all shape names on a slide.
   * @param {number} slideIdx
   * @returns {Array<{id, name, type}>}
   */
  listShapes(slideIdx) {
    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return [];
    const shapes = [];
    for (const child of spTree.children) {
      const ln = child.localName;
      if (!['sp','pic','cxnSp','graphicFrame'].includes(ln)) continue;
      const nvEl  = g1(child, 'nvSpPr') || g1(child, 'nvPicPr') || g1(child, 'nvGraphicFramePr') || g1(child, 'nvCxnSpPr');
      const cNvPr = nvEl ? g1(nvEl, 'cNvPr') : null;
      shapes.push({
        id:   cNvPr?.getAttribute('id') || '',
        name: cNvPr?.getAttribute('name') || '',
        type: ln,
      });
    }
    return shapes;
  }
}
