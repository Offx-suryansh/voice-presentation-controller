/**
 * extract.js — Structured text extraction from PPTX slides.
 *
 * Extracts all readable text content for:
 *   - Search indexing
 *   - Accessibility (screen readers, alt text generation)
 *   - Copy/paste of slide content
 *   - Generating slide outlines / summaries
 *   - Translation pipelines
 *
 * Public API (re-exported from index.js):
 *
 *   renderer.extractSlide(slideIndex)     → SlideContent
 *   renderer.extractAll()                 → SlideContent[]
 *   renderer.extractText(slideIndex)      → string (plain text, all content)
 *   renderer.searchSlides(query)          → SearchResult[]
 *
 * @typedef {object} TextRun
 * @property {string}  text
 * @property {boolean} bold
 * @property {boolean} italic
 * @property {boolean} underline
 * @property {number}  fontSize   pt
 * @property {string}  color      CSS colour string
 *
 * @typedef {object} Paragraph
 * @property {TextRun[]} runs
 * @property {string}    text        — joined plain text
 * @property {string}    align       l | ctr | r | just
 * @property {number}    level       indent level 0–8
 * @property {string|null} bullet    bullet char, number string, or null
 *
 * @typedef {object} TextShape
 * @property {string}      id
 * @property {string}      name
 * @property {string}      type     title | subtitle | body | textBox | other
 * @property {Paragraph[]} paragraphs
 * @property {string}      text     plain text
 *
 * @typedef {object} TableCell
 * @property {number}    row
 * @property {number}    col
 * @property {number}    rowSpan
 * @property {number}    colSpan
 * @property {string}    text
 * @property {Paragraph[]} paragraphs
 *
 * @typedef {object} TableShape
 * @property {string}       id
 * @property {string}       name
 * @property {TableCell[][]} rows
 * @property {string}       text    all cell text joined with tabs/newlines
 *
 * @typedef {object} ImageShape
 * @property {string} id
 * @property {string} name
 * @property {string} altText
 * @property {string} title
 *
 * @typedef {object} ChartShape
 * @property {string}   id
 * @property {string}   name
 * @property {string}   chartType
 * @property {string[]} seriesNames
 * @property {string[]} categories
 *
 * @typedef {object} SlideContent
 * @property {number}       index
 * @property {string}       title        — first title shape text
 * @property {string}       subtitle     — first subtitle shape text
 * @property {TextShape[]}  textShapes
 * @property {TableShape[]} tables
 * @property {ImageShape[]} images
 * @property {ChartShape[]} charts
 * @property {string}       notes        — speaker notes plain text
 * @property {string}       text         — all text joined, for full-text search
 */

// ── XML helpers ───────────────────────────────────────────────────────────────

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
function attrInt(el, name, def = 0) {
  const v = attr(el, name);
  return v !== null ? parseInt(v, 10) : def;
}

// ── Text run extraction ───────────────────────────────────────────────────────

function extractRun(rEl, defRPr) {
  const rPr = g1(rEl, 'rPr');
  const tEl = g1(rEl, 't');
  if (!tEl) return null;

  const text = tEl.textContent || '';
  const szRaw = rPr ? parseInt(rPr.getAttribute('sz') || '0', 10)
    : defRPr ? parseInt(defRPr.getAttribute('sz') || '0', 10) : 0;
  const fontSize = szRaw ? szRaw / 100 : 12; // pt

  const bold   = rPr ? rPr.getAttribute('b') === '1' : false;
  const italic = rPr ? rPr.getAttribute('i') === '1' : false;
  const u      = rPr ? (rPr.getAttribute('u') || 'none') : 'none';

  // Color
  let color = '#000000';
  if (rPr) {
    const solidFill = g1(rPr, 'solidFill') || g1(rPr, 'lumMod');
    if (solidFill) {
      const srgb = g1(solidFill, 'srgbClr');
      if (srgb) color = '#' + (srgb.getAttribute('val') || '000000');
    }
  }

  return { text, bold, italic, underline: u !== 'none', fontSize, color };
}

function extractParagraph(paraEl) {
  const pPr    = g1(paraEl, 'pPr');
  const algn   = attr(pPr, 'algn', 'l');
  const level  = attrInt(pPr, 'lvl', 0);
  const defRPr = g1(pPr, 'defRPr');

  // Bullet
  let bullet = null;
  if (pPr && !g1(pPr, 'buNone')) {
    const buChar    = g1(pPr, 'buChar');
    const buAutoNum = g1(pPr, 'buAutoNum');
    if (buChar)    bullet = buChar.getAttribute('char') || '•';
    else if (buAutoNum) bullet = '{auto}'; // caller can format
  }

  const runs = [];
  for (const child of paraEl.children) {
    if (child.localName === 'r') {
      const run = extractRun(child, defRPr);
      if (run) runs.push(run);
    } else if (child.localName === 'br') {
      runs.push({ text: '\n', bold: false, italic: false, underline: false, fontSize: 12, color: '#000' });
    } else if (child.localName === 'fld') {
      // Field (slide number, date, etc.)
      const t = g1(child, 't');
      if (t) runs.push({ text: t.textContent, bold: false, italic: false, underline: false, fontSize: 12, color: '#555' });
    }
  }

  const text = runs.map(r => r.text).join('');
  return { runs, text, align: algn, level, bullet };
}

function extractTextBody(txBody) {
  if (!txBody) return [];
  return gtn(txBody, 'p').map(extractParagraph).filter(p => p.text.trim());
}

// ── Shape type detection ──────────────────────────────────────────────────────

function detectShapeType(spEl) {
  // Check placeholder type
  const nvSpPr = g1(spEl, 'nvSpPr');
  const nvPr   = nvSpPr ? g1(nvSpPr, 'nvPr') : null;
  const ph     = nvPr ? g1(nvPr, 'ph') : null;

  if (ph) {
    const phType = attr(ph, 'type', 'body');
    if (phType === 'title' || phType === 'ctrTitle') return 'title';
    if (phType === 'subTitle') return 'subtitle';
    if (phType === 'body') return 'body';
  }

  // Check for explicit text box (no placeholder, has txBody)
  const txBody = g1(spEl, 'txBody');
  if (txBody) return 'textBox';

  return 'other';
}

function getShapeId(spEl) {
  const nvSpPr = g1(spEl, 'nvSpPr');
  const cNvPr  = nvSpPr ? g1(nvSpPr, 'cNvPr') : null;
  return cNvPr ? (attr(cNvPr, 'id', '') ) : '';
}

function getShapeName(spEl) {
  const nvSpPr = g1(spEl, 'nvSpPr');
  const cNvPr  = nvSpPr ? g1(nvSpPr, 'cNvPr') : null;
  return cNvPr ? (attr(cNvPr, 'name', '')) : '';
}

function getAltText(spEl) {
  const nvSpPr = g1(spEl, 'nvSpPr');
  const nvPr   = nvSpPr ? g1(nvSpPr, 'nvPr') : null;
  const extLst = nvPr ? g1(nvPr, 'extLst') : null;
  if (extLst) {
    for (const ext of gtn(extLst, 'ext')) {
      const desc = g1(ext, 'lpUserStr') || g1(ext, 'ud');
      if (desc) return desc.getAttribute('val') || desc.textContent || '';
    }
  }
  // title attribute on cNvPr
  const cNvPr = nvSpPr ? g1(nvSpPr, 'cNvPr') : null;
  return cNvPr ? (attr(cNvPr, 'descr', '') || attr(cNvPr, 'title', '')) : '';
}

// ── Table extraction ──────────────────────────────────────────────────────────

function extractTable(graphicFrame) {
  const tbl = g1(graphicFrame, 'tbl');
  if (!tbl) return null;

  const nvGraphicFramePr = g1(graphicFrame, 'nvGraphicFramePr');
  const cNvPr = nvGraphicFramePr ? g1(nvGraphicFramePr, 'cNvPr') : null;
  const id   = cNvPr ? attr(cNvPr, 'id', '') : '';
  const name = cNvPr ? attr(cNvPr, 'name', '') : '';

  const rows = [];
  let ri = 0;
  for (const rowEl of gtn(tbl, 'tr')) {
    const cells = [];
    let ci = 0;
    for (const tcEl of gtn(rowEl, 'tc')) {
      const gridSpan = attrInt(tcEl, 'gridSpan', 1);
      const rowSpan  = attrInt(tcEl, 'rowSpan', 1);
      const paragraphs = extractTextBody(g1(tcEl, 'txBody'));
      const text = paragraphs.map(p => p.text).join('\n');
      cells.push({ row: ri, col: ci, rowSpan, colSpan: gridSpan, text, paragraphs });
      ci++;
    }
    rows.push(cells);
    ri++;
  }

  const text = rows.map(row => row.map(cell => cell.text).join('\t')).join('\n');
  return { id, name, rows, text };
}

// ── Chart extraction ──────────────────────────────────────────────────────────

function extractChartRef(graphicFrame, slideRels) {
  const graphic     = g1(graphicFrame, 'graphic');
  const graphicData = graphic ? g1(graphic, 'graphicData') : null;
  const chartEl     = graphicData ? g1(graphicData, 'chart') : null;
  if (!chartEl) return null;

  const rId = chartEl.getAttribute('r:id') || chartEl.getAttribute('id');
  const nvFramePr = g1(graphicFrame, 'nvGraphicFramePr');
  const cNvPr = nvFramePr ? g1(nvFramePr, 'cNvPr') : null;

  return {
    id:   cNvPr ? attr(cNvPr, 'id', '') : '',
    name: cNvPr ? attr(cNvPr, 'name', '') : '',
    rId,
  };
}

function extractChartContent(chartDoc) {
  if (!chartDoc) return { chartType: 'unknown', seriesNames: [], categories: [] };

  const plotArea = g1(chartDoc, 'plotArea');
  if (!plotArea) return { chartType: 'unknown', seriesNames: [], categories: [] };

  const chartTypes = ['barChart','lineChart','pieChart','areaChart','scatterChart',
                      'doughnutChart','radarChart','bubbleChart','bar3DChart',
                      'line3DChart','pie3DChart','area3DChart'];
  let chartType = 'unknown';
  let chartNode = null;
  for (const t of chartTypes) {
    chartNode = g1(plotArea, t);
    if (chartNode) { chartType = t.replace('3DChart','Chart').replace('Chart',''); break; }
  }

  const serEls = chartNode ? gtn(chartNode, 'ser') : [];
  const seriesNames = serEls.map(s => {
    const tx = g1(s, 'tx');
    if (!tx) return null;
    const v = g1(tx, 'v');
    if (v) return v.textContent.trim();
    const strCache = g1(tx, 'strCache');
    const pt = strCache ? g1(strCache, 'pt') : null;
    const vEl = pt ? g1(pt, 'v') : null;
    return vEl ? vEl.textContent.trim() : null;
  }).filter(Boolean);

  // Categories from first series
  const cats = serEls.length > 0 ? (() => {
    const catEl = g1(serEls[0], 'cat') || g1(serEls[0], 'xVal');
    if (!catEl) return [];
    const cache = g1(catEl, 'strCache') || g1(catEl, 'numCache');
    if (!cache) return [];
    return gtn(cache, 'pt').map(pt => g1(pt, 'v')?.textContent || '').filter(Boolean);
  })() : [];

  return { chartType, seriesNames, categories: cats };
}

// ── Notes extraction ──────────────────────────────────────────────────────────

function extractNotes(notesDoc) {
  if (!notesDoc) return '';
  const cSld  = g1(notesDoc, 'cSld');
  const spTree = cSld ? g1(cSld, 'spTree') : null;
  if (!spTree) return '';

  const parts = [];
  for (const spEl of gtn(spTree, 'sp')) {
    const nvSpPr = g1(spEl, 'nvSpPr');
    const nvPr   = nvSpPr ? g1(nvSpPr, 'nvPr') : null;
    const ph     = nvPr ? g1(nvPr, 'ph') : null;
    // Skip slide number placeholders
    if (ph && attr(ph, 'type') === 'sldNum') continue;
    const txBody = g1(spEl, 'txBody');
    if (txBody) {
      const text = gtn(txBody, 'r')
        .map(r => g1(r, 't')?.textContent || '')
        .join('');
      if (text.trim()) parts.push(text.trim());
    }
  }
  return parts.join('\n\n');
}

// ── Main extractor ─────────────────────────────────────────────────────────────

/**
 * Extract all text content from a single slide.
 *
 * @param {number}    slideIndex
 * @param {object}    renderer    — loaded PptxRenderer instance
 * @returns {Promise<SlideContent>}
 */
export async function extractSlide(slideIndex, renderer) {
  const { _files: files, slidePaths } = renderer;
  if (slideIndex < 0 || slideIndex >= slidePaths.length)
    throw new Error('Slide index out of range');

  const slideXml = files[slidePaths[slideIndex]];
  if (!slideXml) return emptySlide(slideIndex);

  const slideDoc = new DOMParser().parseFromString(
    new TextDecoder().decode(slideXml), 'application/xml'
  );

  // Slide rels (for charts, notes)
  const { getRels } = await import('./render.js');
  const slideRels = await getRels(files, slidePaths[slideIndex]);

  // Notes
  const notesRel = Object.values(slideRels).find(r => r.type?.includes('notesSlide'));
  let notes = '';
  if (notesRel && files[notesRel.fullPath]) {
    const notesDoc = new DOMParser().parseFromString(
      new TextDecoder().decode(files[notesRel.fullPath]), 'application/xml'
    );
    notes = extractNotes(notesDoc);
  }

  const cSld   = g1(slideDoc, 'cSld');
  const spTree = cSld ? g1(cSld, 'spTree') : null;
  if (!spTree) return { index: slideIndex, title: '', subtitle: '', textShapes: [], tables: [], images: [], charts: [], notes, text: notes };

  const textShapes = [];
  const tables     = [];
  const images     = [];
  const charts     = [];

  for (const child of spTree.children) {
    const ln = child.localName;

    if (ln === 'sp') {
      const txBody = g1(child, 'txBody');
      if (!txBody) continue;
      const type       = detectShapeType(child);
      const id         = getShapeId(child);
      const name       = getShapeName(child);
      const paragraphs = extractTextBody(txBody);
      const text       = paragraphs.map(p => p.text).join('\n');
      if (text.trim()) textShapes.push({ id, name, type, paragraphs, text });
    }

    else if (ln === 'pic') {
      const nvPicPr = g1(child, 'nvPicPr');
      const cNvPr   = nvPicPr ? g1(nvPicPr, 'cNvPr') : null;
      const id      = cNvPr ? attr(cNvPr, 'id', '') : '';
      const name    = cNvPr ? attr(cNvPr, 'name', '') : '';
      const altText = cNvPr ? (attr(cNvPr, 'descr', '') || attr(cNvPr, 'title', '')) : '';
      const nvPr    = nvPicPr ? g1(nvPicPr, 'nvPr') : null;
      const cNvPrExt = nvPr ? g1(nvPr, 'extLst') : null;
      images.push({ id, name, altText, title: name });
    }

    else if (ln === 'graphicFrame') {
      const uri = (() => {
        const graphic = g1(child, 'graphic');
        const gd = graphic ? g1(graphic, 'graphicData') : null;
        return gd ? attr(gd, 'uri', '') : '';
      })();

      if (uri.includes('table') || g1(child, 'tbl')) {
        const t = extractTable(child);
        if (t) tables.push(t);
      } else if (uri.includes('chart')) {
        const ref = extractChartRef(child, slideRels);
        if (ref) {
          const rel = ref.rId ? slideRels[ref.rId] : null;
          let chartContent = { chartType: 'chart', seriesNames: [], categories: [] };
          if (rel && files[rel.fullPath]) {
            const chartDoc = new DOMParser().parseFromString(
              new TextDecoder().decode(files[rel.fullPath]), 'application/xml'
            );
            chartContent = extractChartContent(chartDoc);
          }
          charts.push({ id: ref.id, name: ref.name, ...chartContent });
        }
      }
    }

    else if (ln === 'grpSp') {
      // Extract text from group shapes recursively
      for (const spEl of gtn(child, 'sp')) {
        const txBody = g1(spEl, 'txBody');
        if (!txBody) continue;
        const type       = detectShapeType(spEl);
        const id         = getShapeId(spEl);
        const name       = getShapeName(spEl);
        const paragraphs = extractTextBody(txBody);
        const text       = paragraphs.map(p => p.text).join('\n');
        if (text.trim()) textShapes.push({ id, name, type, paragraphs, text });
      }
    }
  }

  // Derive title and subtitle
  const titleShape    = textShapes.find(s => s.type === 'title');
  const subtitleShape = textShapes.find(s => s.type === 'subtitle');
  const title    = titleShape?.text    || '';
  const subtitle = subtitleShape?.text || '';

  // Full text blob
  const allText = [
    title,
    subtitle,
    ...textShapes.filter(s => s.type !== 'title' && s.type !== 'subtitle').map(s => s.text),
    ...tables.map(t => t.text),
    ...charts.map(c => [c.name, ...c.seriesNames, ...c.categories].join(' ')),
    notes,
  ].filter(Boolean).join('\n\n');

  return { index: slideIndex, title, subtitle, textShapes, tables, images, charts, notes, text: allText };
}

function emptySlide(index) {
  return { index, title: '', subtitle: '', textShapes: [], tables: [], images: [], charts: [], notes: '', text: '' };
}

/**
 * Extract content from all slides.
 * @param {object} renderer
 * @returns {Promise<SlideContent[]>}
 */
export async function extractAll(renderer) {
  const results = [];
  for (let i = 0; i < renderer.slideCount; i++) {
    results.push(await extractSlide(i, renderer));
  }
  return results;
}

/**
 * Get all text from a slide as a plain string.
 * @param {number} slideIndex
 * @param {object} renderer
 * @returns {Promise<string>}
 */
export async function extractText(slideIndex, renderer) {
  const content = await extractSlide(slideIndex, renderer);
  return content.text;
}

/**
 * Full-text search across all slides.
 * Case-insensitive, returns slide index + matching excerpts.
 *
 * @param {string} query
 * @param {object} renderer
 * @returns {Promise<SearchResult[]>}
 */
export async function searchSlides(query, renderer) {
  const q = query.toLowerCase().trim();
  if (!q) return [];

  const results = [];
  for (let i = 0; i < renderer.slideCount; i++) {
    const content = await extractSlide(i, renderer);
    const haystack = content.text.toLowerCase();
    if (!haystack.includes(q)) continue;

    // Find all match positions for excerpt generation
    const excerpts = [];
    let pos = 0;
    while ((pos = haystack.indexOf(q, pos)) !== -1) {
      const start  = Math.max(0, pos - 60);
      const end    = Math.min(content.text.length, pos + q.length + 60);
      const before = content.text.slice(start, pos);
      const match  = content.text.slice(pos, pos + q.length);
      const after  = content.text.slice(pos + q.length, end);
      excerpts.push({ before: (start > 0 ? '…' : '') + before, match, after: after + (end < content.text.length ? '…' : '') });
      pos += q.length;
      if (excerpts.length >= 3) break; // max 3 excerpts per slide
    }

    results.push({
      slideIndex: i,
      title: content.title,
      score: excerpts.length + (content.title.toLowerCase().includes(q) ? 10 : 0),
      excerpts,
    });
  }

  return results.sort((a, b) => b.score - a.score);
}
