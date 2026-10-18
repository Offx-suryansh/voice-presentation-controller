/**
 * render.js — Core OOXML rendering: fills, outlines, effects, images,
 * text, shapes, tables, groups, placeholders, background, and the
 * main renderSpTree / renderSlide pipeline.
 *
 * All rendering is done onto a Canvas 2D context.
 */

import { parseXml, g1, gtn, attr, attrInt, EMU_PER_PT } from './utils.js';
import { resolveColorElement, findFirstColorChild, colorToCss, getRunColorInherited } from './colors.js';
import { drawPresetGeom } from './shapes.js';
import { buildFontInherited } from './fonts.js';
import { renderChart } from './charts.js';
import { setup3D, has3D } from './effects3d.js';

// ═══════════════════════════════════════════════════════════════════════════════
//  FILL HELPERS
// ═══════════════════════════════════════════════════════════════════════════════

// Apply a fill to the current path on ctx (assumes path is already set)
export async function applyFill(ctx, fillEl, x, y, w, h, scale, themeColors, imageCache) {
  if (!fillEl) return false;
  const name = fillEl.localName;

  if (name === 'noFill') {
    return false; // no fill
  }

  if (name === 'solidFill') {
    const colorChild = findFirstColorChild(fillEl);
    const c = resolveColorElement(colorChild, themeColors);
    if (c) {
      ctx.fillStyle = colorToCss(c);
      ctx.fill();
      return true;
    }
    return false;
  }

  if (name === 'gradFill') {
    const gsLst = g1(fillEl, 'gsLst');
    if (!gsLst) return false;
    const stops = gtn(gsLst, 'gs').map(gs => {
      const pos = attrInt(gs, 'pos', 0) / 100000;
      const colorChild = findFirstColorChild(gs);
      const c = resolveColorElement(colorChild, themeColors);
      return { pos, color: c };
    }).sort((a, b) => a.pos - b.pos);

    if (stops.length < 2) return false;

    const linEl = g1(fillEl, 'lin');
    const pathEl = g1(fillEl, 'path');

    // cx/cy for this fill region (not the global shape centre — must be defined here)
    const gcx = x + w / 2, gcy = y + h / 2;

    let gradient;
    if (linEl || (!linEl && !pathEl)) {
      // Linear gradient
      // OOXML ang: 0 = top→bottom (north), increases clockwise, units = 60000ths of degree
      const angRaw = attrInt(linEl, 'ang', 0);
      const angDeg = angRaw / 60000;
      // Convert: OOXML 0° = pointing up (north on canvas = -PI/2), clockwise
      const angRad = (angDeg - 90) * Math.PI / 180;
      const cosA = Math.cos(angRad);
      const sinA = Math.sin(angRad);
      // Length of gradient line that covers the full bounding box
      const len = Math.abs(w * cosA) + Math.abs(h * sinA);
      const x1 = gcx - len / 2 * cosA;
      const y1 = gcy - len / 2 * sinA;
      const x2 = gcx + len / 2 * cosA;
      const y2 = gcy + len / 2 * sinA;
      gradient = ctx.createLinearGradient(x1, y1, x2, y2);
    } else {
      // Radial / path gradient
      const fillToRect = g1(pathEl, 'fillToRect');
      // focusL/T/R/B: percentage offsets of the focus rectangle from the shape's edges
      const focusL = attrInt(fillToRect, 'l', 50000) / 100000;
      const focusT = attrInt(fillToRect, 't', 50000) / 100000;
      const focusR = attrInt(fillToRect, 'r', 50000) / 100000;
      const focusB = attrInt(fillToRect, 'b', 50000) / 100000;
      // Focus point = centre of the focus rectangle
      const fx = x + w * (focusL + (1 - focusL - focusR) / 2);
      const fy = y + h * (focusT + (1 - focusT - focusB) / 2);
      // Outer radius: enough to cover corners
      const outerR = Math.sqrt(w * w + h * h) / 2;
      gradient = ctx.createRadialGradient(fx, fy, 0, gcx, gcy, outerR);
    }

    for (const stop of stops) {
      if (stop.color) {
        gradient.addColorStop(stop.pos, colorToCss(stop.color));
      }
    }
    ctx.fillStyle = gradient;
    ctx.fill();
    return true;
  }

  if (name === 'blipFill') {
    const blip = g1(fillEl, 'blip');
    const rEmbed = blip ? (blip.getAttribute('r:embed') || blip.getAttribute('embed')) : null;
    if (rEmbed && imageCache && imageCache[rEmbed]) {
      const img = imageCache[rEmbed];
      const stretch = g1(fillEl, 'stretch');
      const fillRect = stretch ? g1(stretch, 'fillRect') : null;
      let ix = x, iy = y, iw = w, ih = h;
      if (fillRect) {
        const l = attrInt(fillRect, 'l', 0) / 100000;
        const t = attrInt(fillRect, 't', 0) / 100000;
        const r = attrInt(fillRect, 'r', 0) / 100000;
        const b = attrInt(fillRect, 'b', 0) / 100000;
        ix = x + w * l;
        iy = y + h * t;
        iw = w - w * l - w * r;
        ih = h - h * t - h * b;
      }
      // Check for tile
      const tile = g1(fillEl, 'tile');
      if (tile) {
        const pattern = ctx.createPattern(img, 'repeat');
        if (pattern) {
          ctx.fillStyle = pattern;
          ctx.fill();
          return true;
        }
      }
      // Check for alpha/transparency on the blip element
      const alphaMod = g1(fillEl, 'alphaModFix') || g1(blip, 'alphaModFix');
      const alphaVal = alphaMod ? (attrInt(alphaMod, 'amt', 100000) / 100000) : 1;

      ctx.save();
      ctx.clip();
      if (alphaVal < 1) ctx.globalAlpha = alphaVal;
      ctx.drawImage(img, ix, iy, iw, ih);
      ctx.restore();
      return true;
    }
    return false;
  }

  if (name === 'pattFill') {
    // Resolve fg and bg colours
    const fgClrEl = g1(fillEl, 'fgClr');
    const bgClrEl = g1(fillEl, 'bgClr');
    const fgC = fgClrEl ? resolveColorElement(findFirstColorChild(fgClrEl), themeColors) : { r:0,  g:0,  b:0,  a:1 };
    const bgC = bgClrEl ? resolveColorElement(findFirstColorChild(bgClrEl), themeColors) : { r:255,g:255,b:255,a:1 };
    const fgCss = colorToCss(fgC);
    const bgCss = colorToCss(bgC);
    const prst = attr(fillEl, 'prst', 'dotGrid');

    // Build a 4×4 or 8×8 tile on an offscreen canvas
    // Create offscreen tile for pattern — falls back to solid colour in non-browser envs
    let tc = null, tile = null;
    try {
      tile = typeof OffscreenCanvas !== 'undefined'
        ? new OffscreenCanvas(8, 8)
        : document.createElement('canvas');
      const N = 8;
      tile.width = N; tile.height = N;
      tc = tile.getContext('2d');
    } catch (_) {}
    const N = 8;
    if (!tc) {
      // No canvas available — fall back to solid foreground
      ctx.fillStyle = colorToCss(fgC);
      ctx.fill();
      return true;
    }

    // Fill background
    tc.fillStyle = bgCss;
    tc.fillRect(0, 0, N, N);

    // Draw foreground pattern
    tc.fillStyle = fgCss;
    switch (prst) {
      case 'smGrid':
      case 'dotGrid':
      case 'dotDmnd':
        for (let i = 0; i < N; i += 2) {
          for (let j = 0; j < N; j += 2) tc.fillRect(i, j, 1, 1);
        }
        break;
      case 'lgGrid':
      case 'cross':
        tc.fillRect(0, 0, N, 1); tc.fillRect(0, 0, 1, N); // top and left lines
        break;
      case 'diagBd':
      case 'fwdDiag':
        for (let d = 0; d < N * 2; d++) tc.fillRect(d % N, Math.floor(d / N) * 2, 1, 1);
        break;
      case 'bkDiag':
      case 'ltDnDiag':
        for (let d = 0; d < N * 2; d++) tc.fillRect(N - 1 - (d % N), Math.floor(d / N) * 2, 1, 1);
        break;
      case 'horzBrick':
      case 'horz':
        tc.fillRect(0, N / 2, N, 1);
        break;
      case 'vert':
      case 'vertBrick':
        tc.fillRect(N / 2, 0, 1, N);
        break;
      case 'smCheck':
      case 'lgCheck':
        for (let r = 0; r < N; r++) {
          for (let c2 = 0; c2 < N; c2++) {
            if ((r + c2) % 2 === 0) tc.fillRect(c2, r, 1, 1);
          }
        }
        break;
      default:
        // Generic: draw a sparse dot grid
        for (let i = 0; i < N; i += 4) tc.fillRect(i, i, 2, 2);
    }

    const pattern = ctx.createPattern(tile, 'repeat');
    if (pattern) {
      ctx.fillStyle = pattern;
      ctx.fill();
      return true;
    }
    // Fallback: solid foreground colour at reduced opacity
    ctx.fillStyle = colorToCss(fgC, (fgC.a ?? 1) * 0.4);
    ctx.fill();
    return true;
  }

  if (name === 'grpFill') {
    // Group fill - inherits from group, just fill transparent for now
    return false;
  }

  return false;
}
// ═══════════════════════════════════════════════════════════════════════════════
//  RELATIONSHIP RESOLUTION
// ═══════════════════════════════════════════════════════════════════════════════

export async function getRels(files, partPath) {
  // ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels
  const parts = partPath.split('/');
  const filename = parts.pop();
  const relsPath = [...parts, '_rels', filename + '.rels'].join('/');
  const rawData = files[relsPath];
  if (!rawData) return {};
  const content = new TextDecoder().decode(rawData);
  const doc = parseXml(content);
  const rels = {};
  for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    const type = rel.getAttribute('Type') || '';
    const mode = rel.getAttribute('TargetMode') || 'Internal';
    let fullPath = target;
    if (mode !== 'External') {
      if (target.startsWith('/')) {
        fullPath = target.slice(1);
      } else {
        // Resolve relative to the directory of partPath
        const baseParts = partPath.split('/');
        baseParts.pop();
        const targetParts = target.split('/');
        for (const part of targetParts) {
          if (part === '..') baseParts.pop();
          else if (part !== '.') baseParts.push(part);
        }
        fullPath = baseParts.join('/');
      }
    }
    rels[id] = { target, fullPath, type, external: mode === 'External' };
  }
  return rels;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  IMAGE CACHE
// ═══════════════════════════════════════════════════════════════════════════════

export async function loadImages(files, rels) {
  const cache = {};
  const imgExts = new Set(['png','jpg','jpeg','gif','webp','bmp','tiff','tif','svg']);
  for (const [rId, rel] of Object.entries(rels)) {
    if (rel.external) continue;
    const ext = rel.fullPath.split('.').pop().toLowerCase();
    if (!imgExts.has(ext)) continue;
    try {
      const data = files[rel.fullPath];
      if (!data) continue;
      const mimeMap = {
        png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',
        gif:'image/gif',webp:'image/webp',bmp:'image/bmp',
        tiff:'image/tiff',tif:'image/tiff',svg:'image/svg+xml'
      };
      const mime = mimeMap[ext] || 'image/png';
      const blob = new Blob([data], { type: mime });
      const url = URL.createObjectURL(blob);
      const img = await new Promise((resolve, reject) => {
        const image = new Image();
        image.onload = () => resolve(image);
        image.onerror = () => resolve(null);
        image.src = url;
      });
      cache[rId] = img;
    } catch(e) {}
  }
  return cache;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  TEXT RENDERING
// ═══════════════════════════════════════════════════════════════════════════════

// Compute line height from lnSpc element.
// baseSzPx = font size in canvas pixels; scale = EMU → px conversion factor.
export function computeLineHeight(lnSpcEl, baseSzPx, scale) {
  if (!lnSpcEl) return baseSzPx * 1.2;
  const spcPct = g1(lnSpcEl, 'spcPct');
  const spcPts = g1(lnSpcEl, 'spcPts');
  if (spcPct) {
    return baseSzPx * (attrInt(spcPct, 'val', 100000) / 100000);
  } else if (spcPts) {
    // val is in 100ths of a point. 1 pt = 12700 EMU.
    // px = val/100 * 12700 * scale
    const val = attrInt(spcPts, 'val', 0);
    return (val / 100) * 12700 * (scale || 1);
  }
  return baseSzPx * 1.2;
}
// Word-wrap text into lines
/**
 * Wrap text to fit within maxWidth pixels given the current ctx.font.
 * Handles:
 *  - Long words (breaks mid-word if no spaces)
 *  - CJK characters (can wrap between any two characters)
 *  - Trailing spaces preserved per line
 */
export function wrapText(ctx, text, maxWidth) {
  if (maxWidth <= 0 || !text) return text ? [text] : [];

  const CJK_RE = /[\u3000-\u9fff\uac00-\ud7af\uf900-\ufaff\ufe30-\ufeff]/;
  const lines = [];
  let line = '';

  // CJK text: every character is a potential break point, so skip word-splitting
  if (CJK_RE.test(text)) {
    for (const ch of text) {
      const test = line + ch;
      if (ctx.measureText(test).width <= maxWidth) {
        line = test;
      } else {
        if (line) lines.push(line);
        line = ch;
      }
    }
    if (line) lines.push(line);
    return lines.length ? lines : [''];
  }

  const words = text.split(/(\s+)/); // keep whitespace tokens
  for (const token of words) {
    const test = line + token;
    if (ctx.measureText(test).width <= maxWidth) {
      line = test;
    } else if (!line.trim()) {
      // Single long token wider than maxWidth — break by character
      for (const ch of token) {
        const t2 = line + ch;
        if (ctx.measureText(t2).width > maxWidth && line) {
          lines.push(line);
          line = ch;
        } else {
          line = t2;
        }
      }
    } else {
      lines.push(line.trimEnd());
      line = token.trimStart(); // drop leading space at line break
    }
  }
  if (line.trim()) lines.push(line);
  return lines.length ? lines : [''];
}

// Render text body to canvas
// ── BULLET / LIST HELPERS ────────────────────────────────────────────────────

/**
 * Parse bullet properties from a paragraph's <a:pPr> element.
 * Returns null if no bullet, or a bullet descriptor object.
 */
function parseBullet(pPr, defRPr, themeColors, themeData) {
  if (!pPr) return null;

  // Explicit no-bullet
  if (g1(pPr, 'buNone')) return null;

  const buChar    = g1(pPr, 'buChar');
  const buAutoNum = g1(pPr, 'buAutoNum');

  // No bullet element = no bullet (unless inherited, which we skip for now)
  if (!buChar && !buAutoNum) return null;

  // Bullet colour
  let color = null;
  const buClr = g1(pPr, 'buClr');
  if (buClr) {
    const colorChild = findFirstColorChild(buClr);
    color = resolveColorElement(colorChild, themeColors);
  }

  // Bullet size (percentage of run font size)
  const buSzPct = g1(pPr, 'buSzPct');
  const buSzPts = g1(pPr, 'buSzPts');
  let sizePct = 1.0;
  if (buSzPct) sizePct = attrInt(buSzPct, 'val', 100000) / 100000;
  // buSzPts is an absolute size — store as pts for later conversion
  const sizePts = buSzPts ? attrInt(buSzPts, 'val', 0) / 100 : null;

  // Bullet font
  let fontFamily = null;
  const buFont = g1(pPr, 'buFont');
  if (buFont) {
    const tf = buFont.getAttribute('typeface');
    if (tf) fontFamily = tf;
  }

  if (buChar) {
    return {
      type: 'char',
      char: buChar.getAttribute('char') || '•',
      color,
      sizePct,
      sizePts,
      fontFamily,
    };
  }

  if (buAutoNum) {
    return {
      type: 'autoNum',
      numType: buAutoNum.getAttribute('type') || 'arabicPeriod',
      startAt: attrInt(buAutoNum, 'startAt', 1),
      color,
      sizePct,
      sizePts,
      fontFamily,
    };
  }

  return null;
}

/** Auto-number type → formatted string. */
function formatAutoNum(type, n) {
  switch (type) {
    case 'arabicPeriod':    return n + '.';
    case 'arabicParenR':    return n + ')';
    case 'arabicParenBoth': return '(' + n + ')';
    case 'romanLcPeriod':   return toRoman(n).toLowerCase() + '.';
    case 'romanUcPeriod':   return toRoman(n) + '.';
    case 'alphaLcParenR':   return String.fromCharCode(96 + n) + ')';
    case 'alphaUcParenR':   return String.fromCharCode(64 + n) + ')';
    case 'alphaLcPeriod':   return String.fromCharCode(96 + n) + '.';
    case 'alphaUcPeriod':   return String.fromCharCode(64 + n) + '.';
    default:                return n + '.';
  }
}

function toRoman(n) {
  const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
  const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let result = '';
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]) { result += syms[i]; n -= vals[i]; }
  }
  return result;
}

/**
 * Draw a bullet marker at the given position.
 * autoNumCounters: shared map from (numType+startAt) → current count
 */
function drawBullet(ctx, bullet, x, baseline, autoNumCounters) {
  if (!bullet) return;

  // Derive the font size from the current ctx.font (set by the line's first run)
  // We read szPx from ctx.font using a rough parse
  const fontMatch = ctx.font.match(/(d+(?:.d+)?)px/);
  const baseSzPx = fontMatch ? parseFloat(fontMatch[1]) : 16;

  const szPx = bullet.sizePts != null
    ? bullet.sizePts * (baseSzPx / 12) // rough approximation
    : baseSzPx * bullet.sizePct;

  // Save canvas state so bullet doesn't pollute run rendering
  ctx.save();

  // Bullet color (falls back to current fillStyle)
  if (bullet.color) {
    ctx.fillStyle = colorToCss(bullet.color);
    ctx.strokeStyle = ctx.fillStyle;
  }

  // Bullet font
  const family = bullet.fontFamily ? '"' + bullet.fontFamily + '", sans-serif' : ctx.font.split(/\d+px\s+/)[1] || 'sans-serif';
  ctx.font = szPx + 'px ' + family;

  if (bullet.type === 'char') {
    ctx.fillText(bullet.char, x, baseline);
  } else if (bullet.type === 'autoNum') {
    const key = bullet.numType + ':' + bullet.startAt;
    if (autoNumCounters[key] === undefined) autoNumCounters[key] = bullet.startAt;
    const label = formatAutoNum(bullet.numType, autoNumCounters[key]);
    autoNumCounters[key]++;
    ctx.fillText(label, x, baseline);
  }

  ctx.restore();
}

export async function renderTextBody(ctx, txBody, bx, by, bw, bh, scale, themeColors, themeData, defaultFontSz = 1800) {
  if (!txBody) return;

  const bodyPr = g1(txBody, 'bodyPr');
  const anchor = attr(bodyPr, 'anchor', 't'); // t, ctr, b
  const wrap = attr(bodyPr, 'wrap', 'square');
  const vert = attr(bodyPr, 'vert', 'horz'); // horz, vert, vert270, eaVert

  // Text insets (EMU) - OOXML defaults: l=91440, t=45720, r=91440, b=45720
  const lIns = attrInt(bodyPr, 'lIns', 91440) * scale;
  const tIns = attrInt(bodyPr, 'tIns', 45720) * scale;
  const rIns = attrInt(bodyPr, 'rIns', 91440) * scale;
  const bIns = attrInt(bodyPr, 'bIns', 45720) * scale;

  const tx = bx + lIns;
  const ty = by + tIns;
  const tw = bw - lIns - rIns;
  const th = bh - tIns - bIns;

  const doWrap = wrap !== 'none';
  const isVert = vert === 'vert' || vert === 'vert270' || vert === 'eaVert';

  // For vertical text, apply canvas rotation up-front and swap box dimensions
  // so the layout/wrap pass uses the correct axis.
  if (isVert) {
    ctx.save();
    if (vert === 'vert270') {
      ctx.translate(bx + bw, by);
      ctx.rotate(Math.PI / 2);
    } else {
      ctx.translate(bx, by + bh);
      ctx.rotate(-Math.PI / 2);
    }
    // After rotation, height runs along x-axis and width along y-axis — swap them
    // so wrap/layout work in the rotated coordinate space.
    // bx, by, bw, bh are now mapped: new bx=0-relative, bw=original bh, bh=original bw
  }

  // normAutoFit: auto-shrink text to fit. fontScale attr stores the computed scale (0-100000).
  // If no fontScale attr is present, we attempt our own shrink after layout (see second pass).
  const normAutoFit = g1(bodyPr, 'normAutoFit') || g1(txBody, 'normAutoFit');
  const spAutoFit   = g1(bodyPr, 'spAutoFit') || g1(txBody, 'spAutoFit');
  // spAutoFit: box grows to fit text (we can't resize the canvas clip region,
  // but we disable clipping so text is at least visible)
  const explicitFontScale = normAutoFit ? normAutoFit.getAttribute('fontScale') : null;
  let fontScaleAttr = explicitFontScale ? parseInt(explicitFontScale, 10) / 100000 : 1;

  // lstStyle default run properties (lowest priority baseline)
  const lstStyle = g1(txBody, 'lstStyle');
  const lstDefRPr = lstStyle ? g1(lstStyle, 'defRPr') : null;

  // Build a merged "default rPr" helper: returns value from rPr → paraDefRPr → lstDefRPr → fallback
  function resolveRPrAttr(rPr, paraDefRPr, attrName, fallback) {
    const v1 = rPr ? rPr.getAttribute(attrName) : null;
    if (v1 !== null && v1 !== '') return v1;
    const v2 = paraDefRPr ? paraDefRPr.getAttribute(attrName) : null;
    if (v2 !== null && v2 !== '') return v2;
    const v3 = lstDefRPr ? lstDefRPr.getAttribute(attrName) : null;
    if (v3 !== null && v3 !== '') return v3;
    return fallback;
  }

  const paragraphs = gtn(txBody, 'p');

  // First pass: build layout (lines, positions) so we can apply vertical alignment
  const paraLayouts = [];
  let totalHeight = 0;

  for (const para of paragraphs) {
    const pPr = g1(para, 'pPr');
    const algn = attr(pPr, 'algn', 'l'); // l, ctr, r, just, dist
    const lvl = attrInt(pPr, 'lvl', 0);
    const marL = attrInt(pPr, 'marL', 0) * scale;
    const indent = attrInt(pPr, 'indent', 0) * scale;

    // ── Bullet / list marker ─────────────────────────────────────────────────
    const bullet = pPr ? parseBullet(pPr, defRPr, themeColors, themeData) : null;

    // Spacing
    const spcBef = g1(pPr, 'spcBef');
    const spcAft = g1(pPr, 'spcAft');
    const lnSpc = g1(pPr, 'lnSpc');
    const defRPr = g1(pPr, 'defRPr');

    // Default font size: lstStyle (lowest) → pPr defRPr (higher) → run rPr (highest)
    let paraDefSz = defaultFontSz;
    if (lstDefRPr) {
      const sz = lstDefRPr.getAttribute('sz');
      if (sz) paraDefSz = parseInt(sz, 10);
    }
    if (defRPr) {
      const sz = defRPr.getAttribute('sz');
      if (sz) paraDefSz = parseInt(sz, 10);
    }

    // Get space before/after in px
    let spaceBefore = 0, spaceAfter = 0;
    if (spcBef) {
      const sp = g1(spcBef, 'spcPct');
      const spp = g1(spcBef, 'spcPts');
      if (sp) spaceBefore = (paraDefSz * 127 * scale) * (attrInt(sp, 'val', 0) / 100000);
      else if (spp) spaceBefore = attrInt(spp, 'val', 0) * EMU_PER_PT * scale / 100;
    }
    if (spcAft) {
      const sp = g1(spcAft, 'spcPct');
      const spp = g1(spcAft, 'spcPts');
      if (sp) spaceAfter = (paraDefSz * 127 * scale) * (attrInt(sp, 'val', 0) / 100000);
      else if (spp) spaceAfter = attrInt(spp, 'val', 0) * EMU_PER_PT * scale / 100;
    }

    // Collect runs (a:r, a:br)
    const runEls = [];
    for (const child of para.children) {
      const ln = child.localName;
      if (ln === 'r' || ln === 'br' || ln === 'fld') runEls.push(child);
    }

    // Check if paragraph is empty
    if (runEls.length === 0) {
      const endParaRPr = g1(para, 'endParaRPr');
      const sz = attrInt(endParaRPr || defRPr, 'sz', paraDefSz);
      const szPx = sz * 127 * scale * fontScaleAttr;
      paraLayouts.push({ lines: [''], algn, marL, spaceBefore, spaceAfter, szPx, lnSpc, runs: [], emptyPara: true, bullet });
      totalHeight += spaceBefore + szPx * 1.2 + spaceAfter;
      continue;
    }

    // Build text lines by processing runs
    // Each line item: [{ text, rPr, font, color }, ...]
    let paraLines = [];
    let currentLine = [];
    let maxSzPx = 0;

    for (const runEl of runEls) {
      if (runEl.localName === 'br') {
        // Line break
        paraLines.push({ runs: currentLine, maxSzPx: Math.max(maxSzPx, paraDefSz * 127 * scale) });
        currentLine = [];
        maxSzPx = 0;
        continue;
      }

      const rPr = g1(runEl, 'rPr');
      const tEl = g1(runEl, 't');
      let text = tEl ? tEl.textContent : '';

      // Build font using full inheritance chain: rPr → pPr defRPr → lstStyle defRPr
      const fontInfo = buildFontInherited(rPr, defRPr, scale * fontScaleAttr, themeData, paraDefSz, lstDefRPr);
      ctx.font = fontInfo.fontStr;
      const szPx = fontInfo.szPx;
      if (szPx > maxSzPx) maxSzPx = szPx;

      const color = getRunColorInherited(rPr, defRPr, themeColors);
      const underline = resolveRPrAttr(rPr, defRPr, 'u', 'none') !== 'none';
      const strikethrough = resolveRPrAttr(rPr, defRPr, 'strike', 'noStrike') !== 'noStrike';
      const baseline = parseInt(resolveRPrAttr(rPr, defRPr, 'baseline', '0'), 10);

      if (doWrap) {
        // Need to wrap this run's text within remaining line space
        const words = text.split(' ');
        for (let wi = 0; wi < words.length; wi++) {
          const word = words[wi];
          const testRun = { text: word, rPr, fontInfo, color, underline, strikethrough, baseline };
          // Compute current line width
          let lineW = indent + marL;
          for (const run of currentLine) {
            ctx.font = run.fontInfo.fontStr;
            lineW += ctx.measureText(run.text).width;
          }
          ctx.font = fontInfo.fontStr;
          const wordW = ctx.measureText(word).width;
          const sep = currentLine.length ? ctx.measureText(' ').width : 0;

          if (lineW + sep + wordW > tw && currentLine.length > 0) {
            paraLines.push({ runs: currentLine, maxSzPx: Math.max(maxSzPx, szPx) });
            currentLine = [{ text: word, rPr, fontInfo, color, underline, strikethrough, baseline }];
            maxSzPx = szPx;
          } else {
            if (currentLine.length > 0) {
              // Append space to previous run or add space run
              const spaceRun = { text: ' ', rPr, fontInfo, color, underline: false, strikethrough: false, baseline };
              currentLine.push(spaceRun);
            }
            currentLine.push({ text: word, rPr, fontInfo, color, underline, strikethrough, baseline });
          }
        }
      } else {
        currentLine.push({ text, rPr, fontInfo, color, underline, strikethrough, baseline });
        if (szPx > maxSzPx) maxSzPx = szPx;
      }
    }

    if (currentLine.length > 0) {
      paraLines.push({ runs: currentLine, maxSzPx: Math.max(maxSzPx, paraDefSz * 127 * scale) });
    }

    const lnSpcPx = lnSpc ? computeLineHeight(lnSpc, paraDefSz * 127 * scale * fontScaleAttr, scale) : null;
    paraLayouts.push({ lines: paraLines, algn, marL, indent, spaceBefore, spaceAfter, lnSpcPx, emptyPara: false, bullet });

    for (const line of paraLines) {
      totalHeight += spaceBefore + (lnSpcPx || line.maxSzPx * 1.2) + spaceAfter;
    }
  }

  // Auto-shrink: if normAutoFit is present without explicit fontScale,
  // and text overflows the box, iteratively reduce fontScaleAttr until it fits.
  if (normAutoFit && !explicitFontScale && totalHeight > th && th > 0) {
    // Binary-search for a scale that fits (max 8 iterations)
    let lo = 0.3, hi = 1.0;
    for (let iter = 0; iter < 8; iter++) {
      const mid = (lo + hi) / 2;
      // Recompute totalHeight with this scale
      let testH = 0;
      for (const para of paragraphs) {
        const pPr2 = g1(para, 'pPr');
        const defRPr2 = pPr2 ? g1(pPr2, 'defRPr') : null;
        let pSz = defaultFontSz;
        if (lstDefRPr) { const v = lstDefRPr.getAttribute('sz'); if (v) pSz = parseInt(v, 10); }
        if (defRPr2) { const v = defRPr2.getAttribute('sz'); if (v) pSz = parseInt(v, 10); }
        const runEls2 = Array.from(para.children).filter(c => ['r','br','fld'].includes(c.localName));
        const szPx = pSz * 127 * scale * mid;
        if (runEls2.length === 0) { testH += szPx * 1.2; continue; }
        // Estimate: total text length / avg chars per line at this font size
        const totalText = runEls2.reduce((s, e) => {
          const t = g1(e, 't');
          return s + (t ? t.textContent.length : 0);
        }, 0);
        const effectiveTw = tw > 0 ? tw : bw; // fallback to full width if insets eat everything
        // Measure average char width using a representative sample of the actual text
        const sampleText = totalText > 0
          ? runEls2.reduce((s, e) => { const t = g1(e, 't'); return s + (t ? t.textContent : ''); }, '').slice(0, 20)
          : 'W';
        ctx.font = `${szPx}px sans-serif`;
        const avgCharW = sampleText.length > 0 ? ctx.measureText(sampleText).width / sampleText.length : szPx * 0.6;
        const charsPerLine = Math.max(1, Math.floor(effectiveTw / avgCharW));
        const estLines = Math.max(1, Math.ceil(totalText / charsPerLine));
        testH += estLines * szPx * 1.2;
      }
      // Rough estimate: if testH fits, go bigger, else go smaller
      if (testH <= th) lo = mid; else hi = mid;
    }
    fontScaleAttr = (lo + hi) / 2;
  }

  // Vertical alignment
  let startY = ty;
  if (anchor === 'ctr') {
    startY = ty + (th - totalHeight) / 2;
  } else if (anchor === 'b') {
    startY = ty + th - totalHeight;
  }



  // Second pass: render
  let curY = startY;

  // Track auto-numbering counters per bullet type+startAt
  const autoNumCounters = {};

  for (const paraLayout of paraLayouts) {
    const { lines, algn, marL, indent, spaceBefore, spaceAfter, lnSpcPx, emptyPara, bullet } = paraLayout;
    curY += spaceBefore;

    if (emptyPara) {
      curY += lines[0] ? (lnSpcPx || paraLayout.szPx * 1.2) : 12 * scale;
      curY += spaceAfter;
      continue;
    }

    for (const lineObj of lines) {
      const { runs, maxSzPx } = lineObj;
      const lineH = lnSpcPx || maxSzPx * 1.2;
      const baseline = curY + maxSzPx * 0.85; // approximate baseline within line height

      // Calculate total line width for alignment
      let lineW = 0;
      for (const run of runs) {
        ctx.font = run.fontInfo.fontStr;
        lineW += ctx.measureText(run.text).width;
      }

      // ── X-axis start position ─────────────────────────────────────────────
      let runX = tx + marL;
      if (algn === 'ctr') {
        runX = tx + (tw - lineW) / 2;
      } else if (algn === 'r') {
        runX = tx + tw - lineW;
      }

      // ── Bullet / list marker (only on first line of paragraph) ────────────
      const isFirstLineOfPara = lineObj === lines[0];
      if (bullet && isFirstLineOfPara) {
        const bulletX = tx + marL + indent; // indent is typically negative (hanging)
        drawBullet(ctx, bullet, bulletX, baseline, autoNumCounters);
      }

      // ── Justified text: distribute extra space across word gaps ───────────
      let justWordGap = 0;
      if (algn === 'just') {
        const isLastLine = lineObj === lines[lines.length - 1];
        if (!isLastLine) {
          // Count spaces across all runs in this line
          let spaceCount = 0;
          for (const run of runs) {
            ctx.font = run.fontInfo.fontStr;
            spaceCount += (run.text.match(/ /g) || []).length;
          }
          const slack = (tw - marL) - lineW;
          if (spaceCount > 0 && slack > 0) {
            justWordGap = slack / spaceCount;
          }
        }
      }

      // ── Draw a single run at (rx, drawY) — shared by normal + justified paths
      const drawRunSegment = (text, rx, drawY, fi, underline, strike) => {
        // Build font string for this segment (supscript/subscript already handled by caller)
        ctx.font = fi.fontStr;
        const sw = ctx.measureText(text).width;
        ctx.fillText(text, rx, drawY);
        const lw = Math.max(0.5, fi.szPx * 0.07);
        if (underline) {
          ctx.save();
          ctx.strokeStyle = ctx.fillStyle;
          ctx.lineWidth = lw;
          ctx.beginPath();
          ctx.moveTo(rx, drawY + lw * 1.5);
          ctx.lineTo(rx + sw, drawY + lw * 1.5);
          ctx.stroke();
          ctx.restore();
        }
        if (strike) {
          ctx.save();
          ctx.strokeStyle = ctx.fillStyle;
          ctx.lineWidth = lw;
          ctx.beginPath();
          ctx.moveTo(rx, drawY - fi.szPx * 0.3);
          ctx.lineTo(rx + sw, drawY - fi.szPx * 0.3);
          ctx.stroke();
          ctx.restore();
        }
        return sw;
      };

      // ── Draw each run ──────────────────────────────────────────────────────
      for (const run of runs) {
        const c = run.color;
        ctx.fillStyle = c ? colorToCss(c) : '#000000';

        // Superscript / subscript: smaller font + vertical offset
        let drawY = baseline;
        let fi = run.fontInfo;
        if (run.baseline !== 0) {
          const subSz = fi.szPx * 0.65;
          const subFont = `${fi.italic ? 'italic ' : ''}${fi.bold ? 'bold ' : ''}${subSz}px "${fi.family}", sans-serif`;
          fi = { ...fi, szPx: subSz, fontStr: subFont };
          if (run.baseline > 0) drawY = baseline - run.fontInfo.szPx * 0.38; // superscript
          else                   drawY = baseline + run.fontInfo.szPx * 0.12; // subscript
        }

        if (justWordGap > 0 && run.text.includes(' ')) {
          // Justified: render word-by-word with expanded spaces
          ctx.font = fi.fontStr;
          const parts = run.text.split(' ');
          for (let pi = 0; pi < parts.length; pi++) {
            const pw = drawRunSegment(parts[pi], runX, drawY, fi, run.underline, run.strikethrough);
            runX += pw;
            if (pi < parts.length - 1) {
              ctx.font = fi.fontStr;
              runX += ctx.measureText(' ').width + justWordGap;
            }
          }
        } else {
          ctx.font = fi.fontStr;
          const rw = ctx.measureText(run.text).width;
          drawRunSegment(run.text, runX, drawY, fi, run.underline, run.strikethrough);
          runX += rw;
        }
      }

      curY += lineH;
    }
    curY += spaceAfter;
  }

  if (isVert) ctx.restore();
}


// Apply shadow effects from effectLst to the canvas context.
// Must be called BEFORE drawing the shape (sets ctx.shadow*).
// Returns a cleanup function that resets shadow state.
export function applyEffects(ctx, spPr, themeColors, scale) {
  const effectLst = g1(spPr, 'effectLst');
  if (!effectLst) return () => {};

  const outerShdw = g1(effectLst, 'outerShdw');
  const innerShdw = g1(effectLst, 'innerShdw');
  const shadow = outerShdw || innerShdw;

  if (shadow) {
    // blurRad: EMU → pixels
    const blurRad = attrInt(shadow, 'blurRad', 38100) * scale;
    // dist: distance from shape in EMU
    const dist = attrInt(shadow, 'dist', 38100) * scale;
    // dir: angle in 60000ths of degree, clockwise from east
    const dirRaw = attrInt(shadow, 'dir', 2700000);
    const dirRad = dirRaw / 60000 * Math.PI / 180;
    const offsetX = dist * Math.cos(dirRad);
    const offsetY = dist * Math.sin(dirRad);

    // Shadow color
    const colorChild = findFirstColorChild(shadow);
    const c = resolveColorElement(colorChild, themeColors);
    const shadowColor = c ? colorToCss(c) : 'rgba(0,0,0,0.35)';

    ctx.shadowBlur = Math.min(blurRad, 40); // canvas limit ~40px looks good
    ctx.shadowOffsetX = offsetX;
    ctx.shadowOffsetY = offsetY;
    ctx.shadowColor = shadowColor;
  }

  // Glow effect
  const glow = g1(effectLst, 'glow');
  if (glow) {
    const rad = attrInt(glow, 'rad', 0) * scale;
    const colorChild = findFirstColorChild(glow);
    const c = resolveColorElement(colorChild, themeColors);
    if (c) {
      ctx.shadowBlur = Math.min(rad, 30);
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 0;
      ctx.shadowColor = colorToCss(c);
    }
  }

  // Return cleanup
  return () => {
    ctx.shadowBlur = 0;
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;
    ctx.shadowColor = 'transparent';
  };
}

export async function renderShape(ctx, spEl, rels, imageCache, themeColors, themeData, scale, parentGroup = null, placeholderMap = null) {
  const spPr = g1(spEl, 'spPr');
  const xfrm = g1(spPr, 'xfrm');

  // Get bounding box from xfrm
  let x = 0, y = 0, w = 0, h = 0;
  let rot = 0;
  let flipH = false, flipV = false;

  if (xfrm) {
    const off = g1(xfrm, 'off');
    const ext = g1(xfrm, 'ext');
    if (off) {
      x = attrInt(off, 'x', 0) * scale;
      y = attrInt(off, 'y', 0) * scale;
    }
    if (ext) {
      w = attrInt(ext, 'cx', 0) * scale;
      h = attrInt(ext, 'cy', 0) * scale;
    }
    rot = attrInt(xfrm, 'rot', 0) / 60000; // degrees
    flipH = attr(xfrm, 'flipH', '0') === '1';
    flipV = attr(xfrm, 'flipV', '0') === '1';
  } else {
    // No xfrm on the shape — try to resolve from layout/master placeholder map
    const phData = resolvePlaceholderXfrm(spEl, placeholderMap);
    if (phData) {
      x = phData.x * scale;
      y = phData.y * scale;
      w = phData.w * scale;
      h = phData.h * scale;
    } else {
      return; // Can't determine position — skip
    }
  }

  if (w <= 0 || h <= 0) return;

  // Apply group transform adjustment if inside a group
  if (parentGroup) {
    const { grpOff, grpExt, chOff, chExt } = parentGroup;
    const scaleX = grpExt.cx / chExt.cx;
    const scaleY = grpExt.cy / chExt.cy;
    x = grpOff.x + (x / scale - chOff.x) * scaleX * scale;
    y = grpOff.y + (y / scale - chOff.y) * scaleY * scale;
    w = w * scaleX;
    h = h * scaleY;
  }

  const cx = x + w / 2, cy = y + h / 2;

  // Begin drawing with rotation
  ctx.save();
  if (rot !== 0 || flipH || flipV) {
    ctx.translate(cx, cy);
    if (rot !== 0) ctx.rotate(rot * Math.PI / 180);
    if (flipH) ctx.scale(-1, 1);
    if (flipV) ctx.scale(1, -1);
    ctx.translate(-cx, -cy);
  }

  // Get geometry
  const prstGeom = g1(spPr, 'prstGeom');
  const custGeom = g1(spPr, 'custGeom');
  const prst = prstGeom ? attr(prstGeom, 'prst', 'rect') : 'rect';

  // Parse adjustment values
  const adjValues = {};
  if (prstGeom) {
    const avLst = g1(prstGeom, 'avLst');
    if (avLst) {
      let idx = 0;
      for (const gd of gtn(avLst, 'gd')) {
        const fmla = attr(gd, 'fmla', '');
        const m = fmla.match(/val\s+(-?\d+)/);
        if (m) adjValues[idx] = parseInt(m[1]);
        idx++;
      }
    }
  }

  // Draw the path for fill / outline
  const getFill = () => {
    // 1. spPr explicit fill (highest priority)
    const fillNames = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
    for (const fn of fillNames) {
      const el = g1(spPr, fn);
      if (el) return el;
    }
    // 2. style.fillRef — theme-based fill colour for this shape
    const styleEl = g1(spEl, 'style');
    if (styleEl) {
      const fillRef = g1(styleEl, 'fillRef');
      if (fillRef) {
        // fillRef idx=0 means no fill; idx≥1 means use the colour child directly
        const idx = attrInt(fillRef, 'idx', 1);
        if (idx === 0) return parseXml('<noFill/>').documentElement;
        const colorChild = findFirstColorChild(fillRef);
        if (colorChild) {
          // Build a synthetic <solidFill> element in memory using DOMParser
          const ns = 'http://schemas.openxmlformats.org/drawingml/2006/main';
          const doc = parseXml('<solidFill xmlns="' + ns + '">' + colorChild.outerHTML + '</solidFill>');
          return doc.documentElement;
        }
      }
    }
    return null;
  };

  const getOutline = () => {
    // spPr explicit line (highest priority)
    const ln = g1(spPr, 'ln');
    if (ln) return ln;
    // style.lnRef — theme-based line for this shape
    const styleEl = g1(spEl, 'style');
    if (styleEl) {
      const lnRef = g1(styleEl, 'lnRef');
      if (lnRef) {
        const idx = attrInt(lnRef, 'idx', 1);
        if (idx === 0) return null; // explicit no-line
        const colorChild = findFirstColorChild(lnRef);
        if (colorChild) {
          const ns = 'http://schemas.openxmlformats.org/drawingml/2006/main';
          const doc = parseXml('<ln xmlns="' + ns + '"><solidFill>' + colorChild.outerHTML + '</solidFill></ln>');
          return doc.documentElement;
        }
      }
    }
    return null;
  };

  // Handle custom geometry (custGeom) - draw path from pathLst
  if (custGeom) {
    const pathLst = g1(custGeom, 'pathLst');
    if (pathLst) {
      for (const pathEl of gtn(pathLst, 'path')) {
        const pw = attrInt(pathEl, 'w', 1) || 1;
        const ph = attrInt(pathEl, 'h', 1) || 1;
        const sx = w / pw, sy = h / ph;
        ctx.beginPath();
        let cx0 = x, cy0 = y;
        for (const cmd of pathEl.children) {
          switch (cmd.localName) {
            case 'moveTo': {
              const pt = g1(cmd, 'pt');
              if (pt) ctx.moveTo(x + attrInt(pt,'x',0)*sx, y + attrInt(pt,'y',0)*sy);
              break;
            }
            case 'lnTo': {
              const pt = g1(cmd, 'pt');
              if (pt) ctx.lineTo(x + attrInt(pt,'x',0)*sx, y + attrInt(pt,'y',0)*sy);
              break;
            }
            case 'cubicBezTo': {
              const pts = gtn(cmd, 'pt');
              if (pts.length >= 3) {
                ctx.bezierCurveTo(
                  x+attrInt(pts[0],'x',0)*sx, y+attrInt(pts[0],'y',0)*sy,
                  x+attrInt(pts[1],'x',0)*sx, y+attrInt(pts[1],'y',0)*sy,
                  x+attrInt(pts[2],'x',0)*sx, y+attrInt(pts[2],'y',0)*sy
                );
              }
              break;
            }
            case 'quadBezTo': {
              const pts = gtn(cmd, 'pt');
              if (pts.length >= 2) {
                ctx.quadraticCurveTo(
                  x+attrInt(pts[0],'x',0)*sx, y+attrInt(pts[0],'y',0)*sy,
                  x+attrInt(pts[1],'x',0)*sx, y+attrInt(pts[1],'y',0)*sy
                );
              }
              break;
            }
            case 'arcTo': {
              const wR = attrInt(cmd, 'wR', 0)*sx;
              const hR = attrInt(cmd, 'hR', 0)*sy;
              const stAng = attrInt(cmd, 'stAng', 0) / 60000 * Math.PI / 180;
              const swAng = attrInt(cmd, 'swAng', 0) / 60000 * Math.PI / 180;
              // Approximate with canvas arc
              const lastX = ctx._lastX || x;
              const lastY = ctx._lastY || y;
              const ecx = lastX - wR * Math.cos(stAng);
              const ecy = lastY - hR * Math.sin(stAng);
              if (wR === hR) {
                ctx.arc(ecx, ecy, wR, stAng, stAng + swAng, swAng < 0);
              } else {
                ctx.ellipse(ecx, ecy, wR, hR, 0, stAng, stAng + swAng, swAng < 0);
              }
              break;
            }
            case 'close':
              ctx.closePath();
              break;
          }
        }
        // Apply fill/stroke to this custom path
        const fillEl = getFill();
        if (fillEl) await applyFill(ctx, fillEl, x, y, w, h, scale, themeColors, imageCache);
        const lnEl = getOutline();
        if (lnEl) applyOutline(ctx, lnEl, themeColors, scale);
      }
    }
    ctx.restore();
    // Still render text
    const txBody = g1(spEl, 'txBody');
    if (txBody) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, w, h);
      ctx.clip();
      const defSz = getDefaultFontSize(spEl, themeData);
      await renderTextBody(ctx, txBody, x, y, w, h, scale, themeColors, themeData, defSz);
      ctx.restore();
    }
    return;
  }

  // Draw preset or default rect path
  ctx.beginPath();
  const pathDrawn = drawPresetGeom(ctx, prst, x, y, w, h, adjValues);

  // Apply shadow/glow effects before fill (canvas compositing requires this order)
  const cleanupEffects = applyEffects(ctx, spPr, themeColors, scale);

  // 3D: extrusion is drawn before fill (setup3D handles this)
  const fx3d = x, fy3d = y, fw3d = w, fh3d = h;
  const effects3d = has3D(spPr) ? setup3D(ctx, spPr, themeColors, fx3d, fy3d, fw3d, fh3d, scale) : null;
  // camera transform is applied inside setup3D

  try {
    // Apply fill
    const fillEl = getFill();
    let filled = false;

    if (fillEl) {
      filled = await applyFill(ctx, fillEl, x, y, w, h, scale, themeColors, imageCache);
    }

    // 3D overlay (bevel + lighting) goes on top of fill, before outline
    if (effects3d) {
      effects3d.overlay();
    }

    // Apply outline
    const lnEl = getOutline();
    if (lnEl) {
      applyOutline(ctx, lnEl, themeColors, scale);
    } else if (!filled) {
      // No explicit fill or outline — draw a default stroke for line shapes
      if (prst === 'line' || prst === 'straightConnector1') {
        ctx.strokeStyle = '#000';
        ctx.lineWidth = 1;
        ctx.stroke();
      }
    }
  } finally {
    // ALWAYS reset shadow/glow state — guards against early returns and errors
    cleanupEffects();
    if (effects3d?.cleanup) effects3d.cleanup();
  }

  ctx.restore();

  // Render text body (outside rotation - text rotates with shape via ctx transform)
  const txBody = g1(spEl, 'txBody');
  if (txBody) {
    ctx.save();
    if (rot !== 0 || flipH || flipV) {
      ctx.translate(cx, cy);
      if (rot !== 0) ctx.rotate(rot * Math.PI / 180);
      if (flipH) ctx.scale(-1, 1);
      if (flipV) ctx.scale(1, -1);
      ctx.translate(-cx, -cy);
    }
    // Clip to shape bounds
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
    const defSz = getDefaultFontSize(spEl, themeData);
    await renderTextBody(ctx, txBody, x, y, w, h, scale, themeColors, themeData, defSz);
    ctx.restore();
  }
}

export function getDefaultFontSize(spEl, themeData) {
  // Try to get font size hint from placeholder type
  const nvSpPr = g1(spEl, 'nvSpPr');
  const nvPr = nvSpPr ? g1(nvSpPr, 'nvPr') : null;
  const ph = nvPr ? g1(nvPr, 'ph') : null;
  if (ph) {
    const phType = attr(ph, 'type', 'body');
    if (phType === 'title' || phType === 'ctrTitle') return 4400; // 44pt
    if (phType === 'subTitle' || phType === 'body') return 2800; // 28pt
  }
  return 1800; // 18pt default
}

export function applyOutline(ctx, lnEl, themeColors, scale) {
  if (!lnEl) return;
  const noFill = g1(lnEl, 'noFill');
  if (noFill) return;

  const solidFill = g1(lnEl, 'solidFill');
  const gradFill = g1(lnEl, 'gradFill');
  const w = attrInt(lnEl, 'w', 12700); // EMU, default 1pt
  const lineW = Math.max(0.5, w * scale);

  let strokeColor = '#000000';
  if (solidFill) {
    const colorChild = findFirstColorChild(solidFill);
    const c = resolveColorElement(colorChild, themeColors);
    if (c) strokeColor = colorToCss(c);
  } else if (gradFill) {
    strokeColor = '#888888';
  }

  const prstDash = g1(lnEl, 'prstDash');
  const dashType = prstDash ? attr(prstDash, 'val', 'solid') : 'solid';
  const capType = attr(lnEl, 'cap', 'flat');
  const joinType = attr(lnEl, 'cmpd', 'sng');

  ctx.strokeStyle = strokeColor;
  ctx.lineWidth = lineW;
  ctx.lineCap = capType === 'rnd' ? 'round' : capType === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = 'round';

  switch (dashType) {
    case 'dash': ctx.setLineDash([lineW * 4, lineW * 2]); break;
    case 'dot': ctx.setLineDash([lineW, lineW * 2]); break;
    case 'dashDot': ctx.setLineDash([lineW*4, lineW*2, lineW, lineW*2]); break;
    case 'lgDash': ctx.setLineDash([lineW*8, lineW*3]); break;
    case 'lgDashDot': ctx.setLineDash([lineW*8, lineW*3, lineW, lineW*3]); break;
    default: ctx.setLineDash([]);
  }

  ctx.stroke();
  ctx.setLineDash([]);
}

// ═══════════════════════════════════════════════════════════════════════════════
//  PICTURE RENDERING
// ═══════════════════════════════════════════════════════════════════════════════

export async function renderPicture(ctx, picEl, rels, imageCache, themeColors, scale) {
  const spPr = g1(picEl, 'spPr');
  const xfrm = g1(spPr, 'xfrm');
  if (!xfrm) return;

  const off = g1(xfrm, 'off');
  const ext = g1(xfrm, 'ext');
  if (!off || !ext) return;

  const x = attrInt(off, 'x', 0) * scale;
  const y = attrInt(off, 'y', 0) * scale;
  const w = attrInt(ext, 'cx', 0) * scale;
  const h = attrInt(ext, 'cy', 0) * scale;
  const rot = attrInt(xfrm, 'rot', 0) / 60000;
  const flipH = attr(xfrm, 'flipH', '0') === '1';
  const flipV = attr(xfrm, 'flipV', '0') === '1';

  if (w <= 0 || h <= 0) return;

  // Get image reference
  const blipFill = g1(picEl, 'blipFill');
  const blip = blipFill ? g1(blipFill, 'blip') : null;
  const rEmbed = blip ? (blip.getAttribute('r:embed') || blip.getAttribute('embed')) : null;

  const cx = x + w / 2, cy = y + h / 2;

  ctx.save();

  // Apply rotation
  if (rot !== 0 || flipH || flipV) {
    ctx.translate(cx, cy);
    if (rot !== 0) ctx.rotate(rot * Math.PI / 180);
    if (flipH) ctx.scale(-1, 1);
    if (flipV) ctx.scale(1, -1);
    ctx.translate(-cx, -cy);
  }

  if (rEmbed && imageCache[rEmbed]) {
    const img = imageCache[rEmbed];

    // Clipping
    const prstGeom = g1(spPr, 'prstGeom');
    const prst = prstGeom ? attr(prstGeom, 'prst', 'rect') : 'rect';
    const adjValues = {};
    if (prstGeom) {
      const avLst = g1(prstGeom, 'avLst');
      if (avLst) {
        let idx = 0;
        for (const gd of gtn(avLst, 'gd')) {
          const m = (attr(gd, 'fmla', '') || '').match(/val\s+(-?\d+)/);
          if (m) adjValues[idx] = parseInt(m[1]);
          idx++;
        }
      }
    }

    // Clip to shape geometry
    ctx.beginPath();
    drawPresetGeom(ctx, prst, x, y, w, h, adjValues);
    ctx.clip();

    // Determine source crop
    const srcRect = blipFill ? g1(blipFill, 'srcRect') : null;
    if (srcRect) {
      const l = attrInt(srcRect, 'l', 0) / 100000;
      const t = attrInt(srcRect, 't', 0) / 100000;
      const r = attrInt(srcRect, 'r', 0) / 100000;
      const b = attrInt(srcRect, 'b', 0) / 100000;
      const sw = img.naturalWidth * (1 - l - r);
      const sh = img.naturalHeight * (1 - t - b);
      ctx.drawImage(img,
        img.naturalWidth * l, img.naturalHeight * t, sw, sh,
        x, y, w, h);
    } else {
      ctx.drawImage(img, x, y, w, h);
    }

    // Apply outline if any
    const lnEl = g1(spPr, 'ln');
    if (lnEl) {
      ctx.beginPath();
      drawPresetGeom(ctx, prst, x, y, w, h, adjValues);
      applyOutline(ctx, lnEl, themeColors, scale);
    }
  } else {
    // Fallback: draw placeholder rectangle
    ctx.fillStyle = '#e0e0e0';
    ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = '#aaa';
    ctx.lineWidth = 1;
    ctx.strokeRect(x, y, w, h);
    // Draw X
    ctx.beginPath();
    ctx.moveTo(x, y); ctx.lineTo(x+w, y+h);
    ctx.moveTo(x+w, y); ctx.lineTo(x, y+h);
    ctx.stroke();
  }

  ctx.restore();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  TABLE RENDERING
// ═══════════════════════════════════════════════════════════════════════════════

export async function renderTable(ctx, graphicFrame, themeColors, themeData, scale) {
  const xfrm = g1(graphicFrame, 'xfrm');
  if (!xfrm) return;

  const off = g1(xfrm, 'off');
  const ext = g1(xfrm, 'ext');
  if (!off || !ext) return;

  const fx = attrInt(off, 'x', 0) * scale;
  const fy = attrInt(off, 'y', 0) * scale;
  const fw = attrInt(ext, 'cx', 0) * scale;
  const fh = attrInt(ext, 'cy', 0) * scale;

  const graphic = g1(graphicFrame, 'graphic');
  const graphicData = graphic ? g1(graphic, 'graphicData') : null;
  const tbl = graphicData ? g1(graphicData, 'tbl') : null;
  if (!tbl) return;

  const tblPr = g1(tbl, 'tblPr');
  // Band style flags from tblPr
  const bandRow = tblPr ? attr(tblPr, 'bandRow', '0') === '1' : false;
  const bandCol = tblPr ? attr(tblPr, 'bandCol', '0') === '1' : false;
  const firstRow = tblPr ? attr(tblPr, 'firstRow', '0') === '1' : false;
  const lastRow  = tblPr ? attr(tblPr, 'lastRow', '0') === '1' : false;
  const firstCol = tblPr ? attr(tblPr, 'firstCol', '0') === '1' : false;
  const lastCol  = tblPr ? attr(tblPr, 'lastCol', '0') === '1' : false;

  const tblGrid = g1(tbl, 'tblGrid');
  const colWidths = gtn(tblGrid, 'gridCol').map(gc => attrInt(gc, 'w', 0) * scale);

  const rows = gtn(tbl, 'tr');
  let curY = fy;

  for (let ri = 0; ri < rows.length; ri++) {
    const row = rows[ri];
    const rowH = attrInt(row, 'h', 457200) * scale;
    const cells = gtn(row, 'tc');
    let curX = fx;
    const isFirstRow = ri === 0;
    const isLastRow  = ri === rows.length - 1;
    const isOddRow   = ri % 2 === 1;

    for (let ci = 0; ci < cells.length; ci++) {
      const cell = cells[ci];
      const gridSpan = attrInt(cell, 'gridSpan', 1);
      const vMerge = attr(cell, 'vMerge', '0') === '1';

      let cellW = 0;
      for (let gs = 0; gs < gridSpan; gs++) {
        cellW += (colWidths[ci + gs] || 0);
      }

      const tcPr = g1(cell, 'tcPr');
      // Cell fill
      const fillEl = tcPr ? (['noFill','solidFill','gradFill','blipFill','pattFill'].map(n=>g1(tcPr,n)).find(Boolean)) : null;

      ctx.save();
      ctx.beginPath();
      ctx.rect(curX, curY, cellW, rowH);

      if (fillEl) {
        await applyFill(ctx, fillEl, curX, curY, cellW, rowH, scale, themeColors, null);
      } else {
        // Apply band/header coloring from table style flags
        let bandFill = null;
        if (firstRow && isFirstRow) {
          bandFill = themeColors.accent1
            ? '#' + themeColors.accent1.toLowerCase()
            : '#4472C4'; // theme-style header row
        } else if (lastRow && isLastRow) {
          bandFill = '#e0e0e0';
        } else if (bandRow && isOddRow) {
          // Alternating row band — very light tint
          bandFill = 'rgba(0,0,0,0.06)';
        }
        ctx.fillStyle = bandFill || 'transparent';
        if (bandFill) ctx.fill();
      }

      // Cell borders
      const borderProps = [
        { el: g1(tcPr, 'lnL'), x1: curX, y1: curY, x2: curX, y2: curY + rowH },
        { el: g1(tcPr, 'lnR'), x1: curX+cellW, y1: curY, x2: curX+cellW, y2: curY+rowH },
        { el: g1(tcPr, 'lnT'), x1: curX, y1: curY, x2: curX+cellW, y2: curY },
        { el: g1(tcPr, 'lnB'), x1: curX, y1: curY+rowH, x2: curX+cellW, y2: curY+rowH },
      ];

      for (const border of borderProps) {
        if (!border.el) {
          // Default thin border
          ctx.beginPath();
          ctx.strokeStyle = '#cccccc';
          ctx.lineWidth = 0.5;
          ctx.moveTo(border.x1, border.y1);
          ctx.lineTo(border.x2, border.y2);
          ctx.stroke();
        } else {
          const noFill2 = g1(border.el, 'noFill');
          if (!noFill2) {
            ctx.beginPath();
            ctx.moveTo(border.x1, border.y1);
            ctx.lineTo(border.x2, border.y2);
            applyOutline(ctx, border.el, themeColors, scale);
          }
        }
      }

      ctx.restore();

      // Cell text
      const txBody = g1(cell, 'txBody');
      if (txBody) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(curX, curY, cellW, rowH);
        ctx.clip();
        await renderTextBody(ctx, txBody, curX, curY, cellW, rowH, scale, themeColors, themeData, 1400);
        ctx.restore();
      }

      curX += cellW;
    }
    curY += rowH;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  GROUP SHAPE RENDERING
// ═══════════════════════════════════════════════════════════════════════════════

export async function renderGroupShape(ctx, grpSpEl, rels, imageCache, themeColors, themeData, scale) {
  const grpSpPr = g1(grpSpEl, 'grpSpPr');
  const xfrm = g1(grpSpPr, 'xfrm');
  if (!xfrm) return;

  const off = g1(xfrm, 'off');
  const ext = g1(xfrm, 'ext');
  const chOff = g1(xfrm, 'chOff');
  const chExt = g1(xfrm, 'chExt');
  if (!off || !ext || !chOff || !chExt) return;

  const rot = attrInt(xfrm, 'rot', 0) / 60000;
  const flipH = attr(xfrm, 'flipH', '0') === '1';
  const flipV = attr(xfrm, 'flipV', '0') === '1';

  const parentGroup = {
    grpOff: { x: attrInt(off, 'x', 0) * scale, y: attrInt(off, 'y', 0) * scale },
    grpExt: { cx: attrInt(ext, 'cx', 0) * scale, cy: attrInt(ext, 'cy', 0) * scale },
    chOff: { x: attrInt(chOff, 'x', 0), y: attrInt(chOff, 'y', 0) },
    chExt: { cx: attrInt(chExt, 'cx', 1), cy: attrInt(chExt, 'cy', 1) }
  };

  const grpCx = parentGroup.grpOff.x + parentGroup.grpExt.cx / 2;
  const grpCy = parentGroup.grpOff.y + parentGroup.grpExt.cy / 2;

  ctx.save();
  if (rot !== 0 || flipH || flipV) {
    ctx.translate(grpCx, grpCy);
    if (rot !== 0) ctx.rotate(rot * Math.PI / 180);
    if (flipH) ctx.scale(-1, 1);
    if (flipV) ctx.scale(1, -1);
    ctx.translate(-grpCx, -grpCy);
  }

  for (const child of grpSpEl.children) {
    const ln = child.localName;
    if (ln === 'sp') await renderShape(ctx, child, rels, imageCache, themeColors, themeData, scale, parentGroup);
    else if (ln === 'pic') await renderPicture(ctx, child, rels, imageCache, themeColors, scale);
    else if (ln === 'grpSp') await renderGroupShape(ctx, child, rels, imageCache, themeColors, themeData, scale);
    else if (ln === 'graphicFrame') await renderGraphicFrame(ctx, child, themeColors, themeData, scale, files, rels);
    else if (ln === 'cxnSp') await renderConnector(ctx, child, themeColors, scale);
  }
  ctx.restore();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  CONNECTOR RENDERING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Draw an arrowhead or line-end decoration.
 * @param ctx      canvas context
 * @param lnEl     <a:ln> element
 * @param endName  'headEnd' or 'tailEnd'
 * @param tipX,tipY   the pointed tip of the arrow
 * @param fromX,fromY the other end of the line (defines direction)
 */
function drawArrowEnd(ctx, lnEl, endName, tipX, tipY, fromX, fromY, themeColors, scale) {
  const endEl = g1(lnEl, endName);
  if (!endEl) return;
  const type = endEl.getAttribute('type') || 'none';
  if (type === 'none') return;

  // Arrow size from 'w' and 'len' attributes: sm=3, med=6, lg=9 (in lineWidths)
  const sizeMap = { sm: 3, med: 6, lg: 9 };
  const lineW = Math.max(0.5, attrInt(lnEl, 'w', 12700) * scale);
  const aw = lineW * (sizeMap[endEl.getAttribute('w') || 'med'] ?? 6);
  const al = lineW * (sizeMap[endEl.getAttribute('len') || 'med'] ?? 6);

  // Direction angle from tip back toward line
  const angle = Math.atan2(fromY - tipY, fromX - tipX);

  ctx.save();
  ctx.fillStyle = ctx.strokeStyle; // match line colour
  ctx.strokeStyle = ctx.strokeStyle;

  switch (type) {
    case 'triangle':
    case 'arrow':
    case 'stealth': {
      const open = type === 'arrow';
      const indent = type === 'stealth' ? al * 0.5 : 0;
      ctx.beginPath();
      ctx.moveTo(tipX, tipY);
      ctx.lineTo(
        tipX + al * Math.cos(angle) + aw / 2 * Math.sin(angle),
        tipY + al * Math.sin(angle) - aw / 2 * Math.cos(angle)
      );
      if (!open) ctx.lineTo(tipX + indent * Math.cos(angle), tipY + indent * Math.sin(angle));
      ctx.lineTo(
        tipX + al * Math.cos(angle) - aw / 2 * Math.sin(angle),
        tipY + al * Math.sin(angle) + aw / 2 * Math.cos(angle)
      );
      ctx.closePath();
      if (open) ctx.stroke(); else ctx.fill();
      break;
    }
    case 'diamond': {
      const mid = al / 2;
      ctx.beginPath();
      ctx.moveTo(tipX, tipY);
      ctx.lineTo(tipX + mid * Math.cos(angle) + aw/2 * Math.sin(angle),
                 tipY + mid * Math.sin(angle) - aw/2 * Math.cos(angle));
      ctx.lineTo(tipX + al * Math.cos(angle), tipY + al * Math.sin(angle));
      ctx.lineTo(tipX + mid * Math.cos(angle) - aw/2 * Math.sin(angle),
                 tipY + mid * Math.sin(angle) + aw/2 * Math.cos(angle));
      ctx.closePath();
      ctx.fill();
      break;
    }
    case 'oval': {
      const rx = al / 2, ry = aw / 2;
      const cx = tipX + rx * Math.cos(angle), cy = tipY + rx * Math.sin(angle);
      ctx.beginPath();
      ctx.save();
      ctx.translate(cx, cy);
      ctx.rotate(angle);
      ctx.scale(1, ry / rx);
      ctx.arc(0, 0, rx, 0, Math.PI * 2);
      ctx.restore();
      ctx.fill();
      break;
    }
  }
  ctx.restore();
}

export async function renderConnector(ctx, cxnSpEl, themeColors, scale) {
  const spPr = g1(cxnSpEl, 'spPr');
  const xfrm = g1(spPr, 'xfrm');
  if (!xfrm) return;

  const off = g1(xfrm, 'off');
  const ext = g1(xfrm, 'ext');
  if (!off || !ext) return;

  const x = attrInt(off, 'x', 0) * scale;
  const y = attrInt(off, 'y', 0) * scale;
  const w = attrInt(ext, 'cx', 0) * scale;
  const h = attrInt(ext, 'cy', 0) * scale;
  const rot = attrInt(xfrm, 'rot', 0) / 60000;
  const flipH = attr(xfrm, 'flipH', '0') === '1';
  const flipV = attr(xfrm, 'flipV', '0') === '1';
  const cx = x + w/2, cy = y + h/2;

  const lnEl = g1(spPr, 'ln');

  ctx.save();
  if (rot !== 0 || flipH || flipV) {
    ctx.translate(cx, cy);
    if (rot !== 0) ctx.rotate(rot * Math.PI / 180);
    if (flipH) ctx.scale(-1, 1);
    if (flipV) ctx.scale(1, -1);
    ctx.translate(-cx, -cy);
  }

  const prstGeom = g1(spPr, 'prstGeom');
  const prst = prstGeom ? attr(prstGeom, 'prst', 'line') : 'line';

  // For line connectors, draw the actual line and arrowheads explicitly
  // (avoids canvas path coordinate limitations)
  const isLine = prst === 'line' || prst === 'straightConnector1';

  ctx.beginPath();
  if (isLine) {
    const x2 = flipH ? x : x + w;
    const y2 = flipV ? y : y + h;
    const x1 = flipH ? x + w : x;
    const y1 = flipV ? y + h : y;
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
  } else {
    drawPresetGeom(ctx, prst, x, y, w, h, {});
  }

  if (lnEl) {
    applyOutline(ctx, lnEl, themeColors, scale);
    // Draw arrowheads for straight lines
    if (isLine) {
      const x1r = flipH ? x + w : x, y1r = flipV ? y + h : y;
      const x2r = flipH ? x : x + w, y2r = flipV ? y : y + h;
      drawArrowEnd(ctx, lnEl, 'headEnd', x2r, y2r, x1r, y1r, themeColors, scale);
      drawArrowEnd(ctx, lnEl, 'tailEnd', x1r, y1r, x2r, y2r, themeColors, scale);
    }
  } else {
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 1;
    ctx.stroke();
  }

  ctx.restore();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  BACKGROUND RENDERING
// ═══════════════════════════════════════════════════════════════════════════════

export async function renderBackground(ctx, slideDoc, masterDoc, layoutDoc, rels, masterRels, imageCache, themeColors, scale, slideW, slideH) {
  const canvasW = slideW * scale;
  const canvasH = slideH * scale;

  // Try to get background from slide, then layout, then master
  const getBg = (doc) => {
    const cSld = g1(doc, 'cSld');
    if (!cSld) return null;
    const bg = g1(cSld, 'bg');
    if (!bg) return null;
    const bgPr = g1(bg, 'bgPr');
    const bgRef = g1(bg, 'bgRef');
    return { bgPr, bgRef };
  };

  const slideBg = slideDoc ? getBg(slideDoc) : null;
  const layoutBg = layoutDoc ? getBg(layoutDoc) : null;
  const masterBg = masterDoc ? getBg(masterDoc) : null;

  const bgData = slideBg || layoutBg || masterBg;

  let rendered = false;
  if (bgData) {
    const { bgPr, bgRef } = bgData;
    if (bgPr) {
      const fills = ['noFill','solidFill','gradFill','blipFill','pattFill'];
      for (const fn of fills) {
        const fillEl = g1(bgPr, fn);
        if (fillEl) {
          ctx.beginPath();
          ctx.rect(0, 0, canvasW, canvasH);
          const useCache = bgData === masterBg ? Object.assign({}, imageCache) : imageCache;
          const ok = await applyFill(ctx, fillEl, 0, 0, canvasW, canvasH, 1, themeColors, useCache);
          if (ok) rendered = true;
          break;
        }
      }
    } else if (bgRef) {
      const idx = attrInt(bgRef, 'idx', 0);
      const colorChild = findFirstColorChild(bgRef);
      const c = resolveColorElement(colorChild, themeColors);
      if (c) {
        ctx.fillStyle = colorToCss(c);
        ctx.fillRect(0, 0, canvasW, canvasH);
        rendered = true;
      }
    }
  }

  if (!rendered) {
    // Default white background
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvasW, canvasH);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  PLACEHOLDER INHERITANCE
//  Build a map of ph type/idx → { x, y, w, h, txBody } from layout & master.
//  Used when a slide's placeholder shape has no xfrm of its own.
// ═══════════════════════════════════════════════════════════════════════════════

export function buildPlaceholderMap(docs) {
  // docs: [layoutDoc, masterDoc] — layout takes priority
  const map = {};
  for (const doc of docs) {
    if (!doc) continue;
    const cSld = g1(doc, 'cSld');
    const spTree = cSld ? g1(cSld, 'spTree') : null;
    if (!spTree) continue;
    for (const sp of gtn(spTree, 'sp')) {
      const nvSpPr = g1(sp, 'nvSpPr');
      const nvPr = nvSpPr ? g1(nvSpPr, 'nvPr') : null;
      const ph = nvPr ? g1(nvPr, 'ph') : null;
      if (!ph) continue;

      const phType = attr(ph, 'type', 'body');
      const phIdx = attr(ph, 'idx', '0');
      const key = `${phType}:${phIdx}`;

      if (map[key]) continue; // layout wins over master

      const spPr = g1(sp, 'spPr');
      const xfrm = g1(spPr, 'xfrm');
      if (!xfrm) continue;
      const off = g1(xfrm, 'off');
      const ext = g1(xfrm, 'ext');
      if (!off || !ext) continue;

      map[key] = {
        x: attrInt(off, 'x', 0),
        y: attrInt(off, 'y', 0),
        w: attrInt(ext, 'cx', 0),
        h: attrInt(ext, 'cy', 0),
        txBody: g1(sp, 'txBody'),
      };
    }
  }
  return map;
}

// Look up placeholder position for a slide shape that has no xfrm
export function resolvePlaceholderXfrm(spEl, placeholderMap) {
  if (!placeholderMap) return null;
  const nvSpPr = g1(spEl, 'nvSpPr');
  const nvPr = nvSpPr ? g1(nvSpPr, 'nvPr') : null;
  const ph = nvPr ? g1(nvPr, 'ph') : null;
  if (!ph) return null;
  const phType = attr(ph, 'type', 'body');
  const phIdx = attr(ph, 'idx', '0');
  // Try exact key, then idx-only, then type-only
  return placeholderMap[`${phType}:${phIdx}`]
      || placeholderMap[`${phType}:0`]
      || placeholderMap[`body:${phIdx}`]
      || null;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  CHART PLACEHOLDER
// ═══════════════════════════════════════════════════════════════════════════════

export async function renderGraphicFrame(ctx, graphicFrame, themeColors, themeData, scale, files, slideRels) {
  const graphic = g1(graphicFrame, 'graphic');
  const graphicData = graphic ? g1(graphic, 'graphicData') : null;
  const uri = graphicData ? attr(graphicData, 'uri', '') : '';

  // Table
  if (g1(graphicFrame, 'tbl') || (graphicData && g1(graphicData, 'tbl'))) {
    return renderTable(ctx, graphicFrame, themeColors, themeData, scale);
  }

  const xfrm = g1(graphicFrame, 'xfrm');
  if (!xfrm) return;
  const off = g1(xfrm, 'off'), ext = g1(xfrm, 'ext');
  if (!off || !ext) return;
  const fx = attrInt(off, 'x', 0) * scale;
  const fy = attrInt(off, 'y', 0) * scale;
  const fw = attrInt(ext, 'cx', 0) * scale;
  const fh = attrInt(ext, 'cy', 0) * scale;
  if (fw <= 0 || fh <= 0) return;

  const isChart   = uri.includes('chart');
  const isDiagram = uri.includes('diagram');

  // ── Real chart rendering ─────────────────────────────────────────────────
  if (isChart && files && slideRels) {
    // Find the chart relationship
    const chartEl = graphicData ? g1(graphicData, 'chart') : null;
    const rId = chartEl
      ? (chartEl.getAttribute('r:id') || chartEl.getAttribute('id'))
      : null;
    const rel = rId ? slideRels[rId] : null;

    if (rel && files[rel.fullPath]) {
      const chartXml = new TextDecoder().decode(files[rel.fullPath]);
      const chartDoc = parseXml(chartXml);
      renderChart(ctx, chartDoc, fx, fy, fw, fh, themeColors, scale);
      return;
    }
  }

  // ── SmartArt / Diagram rendering ─────────────────────────────────────────
  if (isDiagram && files && slideRels) {
    const dgmEl = graphicData ? g1(graphicData, 'relIds') : null;
    const dmId  = dgmEl ? (dgmEl.getAttribute('r:dm') || dgmEl.getAttribute('dm')) : null;
    const rel   = dmId ? slideRels[dmId] : null;
    if (rel && files[rel.fullPath]) {
      const dataXml = new TextDecoder().decode(files[rel.fullPath]);
      const dataDoc = parseXml(dataXml);
      // Try to find a layout file
      const loId  = dgmEl ? (dgmEl.getAttribute('r:lo') || dgmEl.getAttribute('lo')) : null;
      const loRel = loId ? slideRels[loId] : null;
      const layoutDoc = (loRel && files[loRel.fullPath])
        ? parseXml(new TextDecoder().decode(files[loRel.fullPath]))
        : null;
      // Delegate to SmartArt renderer (imported lazily to keep this file lean)
      const { renderSmartArt } = await import('./smartart.js');
      renderSmartArt(ctx, dataDoc, layoutDoc, fx, fy, fw, fh, themeColors, scale);
      return;
    }
  }

  // ── Fallback placeholder ─────────────────────────────────────────────────
  const label = isChart ? '📊 Chart' : isDiagram ? '🔷 Diagram' : '⬛ Graphic';
  ctx.save();
  ctx.fillStyle = '#f4f4f8';
  ctx.strokeStyle = '#ccccdd';
  ctx.lineWidth = 1;
  ctx.fillRect(fx, fy, fw, fh);
  ctx.strokeRect(fx, fy, fw, fh);
  ctx.fillStyle = '#999';
  ctx.font = `${Math.min(fw * 0.07, 16 * scale)}px sans-serif`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(label, fx + fw / 2, fy + fh / 2);
  ctx.textAlign = 'start';
  ctx.textBaseline = 'alphabetic';
  ctx.restore();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SHAPE TREE RENDERING
// ═══════════════════════════════════════════════════════════════════════════════

export async function renderSpTree(ctx, spTreeEl, rels, imageCache, themeColors, themeData, scale, placeholderMap, files) {
  if (!spTreeEl) return;
  for (const child of spTreeEl.children) {
    const ln = child.localName;
    try {
      if (ln === 'sp') await renderShape(ctx, child, rels, imageCache, themeColors, themeData, scale, null, placeholderMap);
      else if (ln === 'pic') await renderPicture(ctx, child, rels, imageCache, themeColors, scale);
      else if (ln === 'grpSp') await renderGroupShape(ctx, child, rels, imageCache, themeColors, themeData, scale);
      else if (ln === 'graphicFrame') await renderGraphicFrame(ctx, child, themeColors, themeData, scale, files, rels);
      else if (ln === 'cxnSp') await renderConnector(ctx, child, themeColors, scale);
    } catch(e) {
      console.warn('Error rendering shape:', ln, e);
    }
  }
}
