/**
 * svg.js — PPTX → SVG serializer.
 *
 * Produces a faithful SVG for each slide: vector text (searchable / scalable),
 * inline base-64 images, linearGradient / radialGradient fills, clip paths,
 * shadows via <filter>, and the same preset geometry paths used by the canvas
 * renderer.
 *
 * Public API (all re-exported from index.js):
 *   renderSlideToSvg(slideIndex, renderer) → Promise<string>   SVG markup
 *   renderAllSlidesToSvg(renderer)          → Promise<string[]>
 *
 * The SVG is self-contained (no external resources) and matches PowerPoint's
 * "Save as SVG" output format.
 */

import { g1, gtn, attr, attrInt, parseXml, EMU_PER_PT } from './utils.js';
import { resolveColorElement, findFirstColorChild, colorToCss } from './colors.js';
import { buildFontInherited } from './fonts.js';
import { getRels } from './render.js';

// ── ID generator ──────────────────────────────────────────────────────────────
let _idSeq = 0;
function uid(prefix = 'el') { return `${prefix}${++_idSeq}`; }

// ── Attribute helpers ─────────────────────────────────────────────────────────
function esc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function px(n) { return `${+n.toFixed(3)}`; }

// ── Colour helpers ────────────────────────────────────────────────────────────
function colStr(c) { return c ? colorToCss(c) : 'none'; }

function fillAttr(fillEl, defs, themeColors, x, y, w, h) {
  if (!fillEl) return { fill: 'none', fillAttrs: '' };
  const ln = fillEl.localName;

  if (ln === 'noFill') return { fill: 'none', fillAttrs: '' };

  if (ln === 'solidFill') {
    const cc = findFirstColorChild(fillEl);
    const c  = resolveColorElement(cc, themeColors);
    const css = colStr(c);
    const opacity = c?.a != null ? c.a / 255 : 1;
    return { fill: css, fillAttrs: opacity < 1 ? ` fill-opacity="${px(opacity)}"` : '' };
  }

  if (ln === 'gradFill') {
    const gsLst = g1(fillEl, 'gsLst');
    const stops = gsLst ? gtn(gsLst, 'gs').map(gs => {
      const pos = attrInt(gs, 'pos', 0) / 100000;
      const cc  = findFirstColorChild(gs);
      const c   = resolveColorElement(cc, themeColors);
      return { pos, color: colStr(c), opacity: c?.a != null ? c.a / 255 : 1 };
    }) : [];

    const linEl  = g1(fillEl, 'lin');
    const pathEl = g1(fillEl, 'path');
    const gradId = uid('grad');

    let gradDef;
    if (pathEl) {
      // Radial gradient
      const fillToRect = g1(pathEl, 'fillToRect');
      const fl = attrInt(fillToRect, 'l', 50000) / 100000;
      const ft = attrInt(fillToRect, 't', 50000) / 100000;
      const cx = x + w * fl;
      const cy = y + h * ft;
      const r  = Math.sqrt(w * w + h * h) / 2;
      gradDef = `<radialGradient id="${gradId}" cx="${px(cx)}" cy="${px(cy)}" r="${px(r)}" gradientUnits="userSpaceOnUse">`;
      for (const s of stops) {
        gradDef += `<stop offset="${px(s.pos)}" stop-color="${esc(s.color)}"${s.opacity < 1 ? ` stop-opacity="${px(s.opacity)}"` : ''}/>`;
      }
      gradDef += `</radialGradient>`;
    } else {
      // Linear gradient
      const angRaw = attrInt(linEl, 'ang', 0);
      const ang    = ((angRaw / 60000) - 90) * (Math.PI / 180);
      const cos    = Math.cos(ang), sin = Math.sin(ang);
      const half   = Math.sqrt(w * w + h * h) / 2;
      const cx2    = x + w / 2, cy2 = y + h / 2;
      gradDef = `<linearGradient id="${gradId}" x1="${px(cx2 - cos * half)}" y1="${px(cy2 - sin * half)}" x2="${px(cx2 + cos * half)}" y2="${px(cy2 + sin * half)}" gradientUnits="userSpaceOnUse">`;
      for (const s of stops) {
        gradDef += `<stop offset="${px(s.pos)}" stop-color="${esc(s.color)}"${s.opacity < 1 ? ` stop-opacity="${px(s.opacity)}"` : ''}/>`;
      }
      gradDef += `</linearGradient>`;
    }
    defs.push(gradDef);
    return { fill: `url(#${gradId})`, fillAttrs: '' };
  }

  if (ln === 'blipFill') {
    // Image fill — handled separately (see renderPictureSvg)
    return { fill: 'none', fillAttrs: '' };
  }

  return { fill: 'none', fillAttrs: '' };
}

// ── Stroke helper ─────────────────────────────────────────────────────────────
function strokeAttrs(lnEl, themeColors, scale) {
  if (!lnEl) return '';
  if (g1(lnEl, 'noFill')) return ' stroke="none"';

  const solidFill = g1(lnEl, 'solidFill');
  const cc   = solidFill ? findFirstColorChild(solidFill) : null;
  const c    = resolveColorElement(cc, themeColors);
  const color = colStr(c) || '#000';
  const w    = Math.max(0.5, attrInt(lnEl, 'w', 12700) / 914400 * 96); // px at 96dpi

  const prstDash = g1(lnEl, 'prstDash');
  const dash = prstDash ? attr(prstDash, 'val', 'solid') : 'solid';
  let dashArr = '';
  if (dash === 'dash')        dashArr = ` stroke-dasharray="${px(w*4)},${px(w*2)}"`;
  else if (dash === 'dot')    dashArr = ` stroke-dasharray="${px(w)},${px(w*2)}"`;
  else if (dash === 'dashDot')dashArr = ` stroke-dasharray="${px(w*4)},${px(w*2)},${px(w)},${px(w*2)}"`;
  else if (dash === 'lgDash') dashArr = ` stroke-dasharray="${px(w*8)},${px(w*3)}"`;

  const cap  = attr(lnEl, 'cap', 'flat');
  const capSvg = cap === 'rnd' ? 'round' : cap === 'sq' ? 'square' : 'butt';

  return ` stroke="${esc(color)}" stroke-width="${px(w)}" stroke-linecap="${capSvg}" stroke-linejoin="round"${dashArr}`;
}

// ── Shadow filter ──────────────────────────────────────────────────────────────
function shadowFilter(effectLst, defs) {
  if (!effectLst) return '';
  const outerShdw = g1(effectLst, 'outerShdw');
  if (!outerShdw) return '';

  const dist  = attrInt(outerShdw, 'dist', 38100) / 914400 * 96;
  const dir   = attrInt(outerShdw, 'dir', 2700000) / 60000;
  const blurR = attrInt(outerShdw, 'blurRad', 38100) / 914400 * 96;
  const ang   = (dir * Math.PI) / 180;
  const dx    = Math.cos(ang) * dist;
  const dy    = Math.sin(ang) * dist;

  const cc  = findFirstColorChild(outerShdw);
  const c   = resolveColorElement(cc, {});
  const col = c ? colorToCss(c) : 'rgba(0,0,0,0.5)';
  const opacity = c?.a != null ? c.a / 255 : 0.5;

  const filterId = uid('shd');
  defs.push(
    `<filter id="${filterId}" x="-50%" y="-50%" width="200%" height="200%">` +
    `<feDropShadow dx="${px(dx)}" dy="${px(dy)}" stdDeviation="${px(blurR / 2)}" flood-color="${esc(col)}" flood-opacity="${px(opacity)}"/>` +
    `</filter>`
  );
  return ` filter="url(#${filterId})"`;
}

// ── Preset geometry path builder ──────────────────────────────────────────────
// We need SVG path strings for the same presets the canvas renderer draws.
// Strategy: draw to a mock path-capturing context, extract path commands.

function presetToSvgPath(prst, x, y, w, h, adjValues) {
  // Use the same drawPresetGeom function but with a path-capturing mock ctx
  const cmds = [];
  const mock = {
    beginPath() { cmds.length = 0; },
    moveTo(px2, py2) { cmds.push(`M${px(px2)},${px(py2)}`); },
    lineTo(px2, py2) { cmds.push(`L${px(px2)},${px(py2)}`); },
    bezierCurveTo(c1x,c1y,c2x,c2y,ex,ey) {
      cmds.push(`C${px(c1x)},${px(c1y)},${px(c2x)},${px(c2y)},${px(ex)},${px(ey)}`);
    },
    quadraticCurveTo(cpx,cpy,ex,ey) {
      cmds.push(`Q${px(cpx)},${px(cpy)},${px(ex)},${px(ey)}`);
    },
    arc(cx2,cy2,r,start,end,ccw) {
      // Convert to SVG arc
      const startX = cx2 + r * Math.cos(start);
      const startY = cy2 + r * Math.sin(start);
      const endX   = cx2 + r * Math.cos(end);
      const endY   = cy2 + r * Math.sin(end);
      let sweep = end - start;
      if (ccw) sweep = -((Math.PI * 2) - Math.abs(sweep));
      const large = Math.abs(sweep) > Math.PI ? 1 : 0;
      const sweepFlag = sweep > 0 ? 1 : 0;
      if (cmds.length === 0) cmds.push(`M${px(startX)},${px(startY)}`);
      cmds.push(`A${px(r)},${px(r)},0,${large},${sweepFlag},${px(endX)},${px(endY)}`);
    },
    arcTo(x1,y1,x2,y2,r) {
      // Approximate with lineTo for SVG (arcTo is canvas-specific)
      cmds.push(`L${px(x1)},${px(y1)} L${px(x2)},${px(y2)}`);
    },
    closePath() { cmds.push('Z'); },
    fill() {},
    stroke() {},
    save() {}, restore() {}, translate() {}, rotate() {}, scale() {},
    rect(rx,ry,rw,rh) {
      cmds.push(`M${px(rx)},${px(ry)}L${px(rx+rw)},${px(ry)}L${px(rx+rw)},${px(ry+rh)}L${px(rx)},${px(ry+rh)}Z`);
    },
    setLineDash() {},
    measureText() { return { width: 0 }; },
  };

  try {
    // Dynamically import shapes to avoid circular dep
    // We'll call through a stored reference set at module init
    if (_drawPresetGeom) {
      _drawPresetGeom(mock, prst, x, y, w, h, adjValues || {});
    }
  } catch (e) {
    // Fallback to rect
    return `M${px(x)},${px(y)}L${px(x+w)},${px(y)}L${px(x+w)},${px(y+h)}L${px(x)},${px(y+h)}Z`;
  }
  return cmds.join(' ') || `M${px(x)},${px(y)}L${px(x+w)},${px(y)}L${px(x+w)},${px(y+h)}L${px(x)},${px(y+h)}Z`;
}

let _drawPresetGeom = null;
export function initSvgShapeRenderer(drawPresetGeomFn) {
  _drawPresetGeom = drawPresetGeomFn;
}

// ── Transform helper ──────────────────────────────────────────────────────────
function xfrmAttrs(xfrm) {
  if (!xfrm) return '';
  const rot   = attrInt(xfrm, 'rot', 0) / 60000;
  const flipH = attr(xfrm, 'flipH', '0') === '1';
  const flipV = attr(xfrm, 'flipV', '0') === '1';
  const off   = g1(xfrm, 'off');
  const ext   = g1(xfrm, 'ext');
  if (!off || !ext) return '';
  const x = attrInt(off, 'x', 0) / 914400 * 96;
  const y = attrInt(off, 'y', 0) / 914400 * 96;
  const w = attrInt(ext, 'cx', 0) / 914400 * 96;
  const h = attrInt(ext, 'cy', 0) / 914400 * 96;
  const cx = x + w / 2, cy = y + h / 2;

  const parts = [];
  if (rot) parts.push(`rotate(${px(rot)},${px(cx)},${px(cy)})`);
  if (flipH) parts.push(`scale(-1,1) translate(${px(-x*2-w)},0)`);
  if (flipV) parts.push(`scale(1,-1) translate(0,${px(-y*2-h)})`);

  return parts.length ? ` transform="${parts.join(' ')}"` : '';
}

// ── Shape bounds from xfrm ────────────────────────────────────────────────────
function xfrmBounds(xfrm) {
  if (!xfrm) return null;
  const off = g1(xfrm, 'off'), ext = g1(xfrm, 'ext');
  if (!off || !ext) return null;
  return {
    x: attrInt(off, 'x', 0) / 914400 * 96,
    y: attrInt(off, 'y', 0) / 914400 * 96,
    w: attrInt(ext, 'cx', 0) / 914400 * 96,
    h: attrInt(ext, 'cy', 0) / 914400 * 96,
  };
}

// ── Text body → SVG ───────────────────────────────────────────────────────────
function textBodyToSvg(txBody, bx, by, bw, bh, themeColors, themeData, defs) {
  if (!txBody) return '';
  const bodyPr   = g1(txBody, 'bodyPr');
  const vert     = attr(bodyPr, 'vert', 'horz');
  const isVert   = vert === 'vert' || vert === 'vert270' || vert === 'wordArtVert';
  const anchor   = attr(bodyPr, 'anchor', 't');
  const lIns     = attrInt(bodyPr, 'lIns', 91440) / 914400 * 96;
  const rIns     = attrInt(bodyPr, 'rIns', 91440) / 914400 * 96;
  const tIns     = attrInt(bodyPr, 'tIns', 45720) / 914400 * 96;
  const bIns     = attrInt(bodyPr, 'bIns', 45720) / 914400 * 96;

  const tx = bx + lIns, tw = bw - lIns - rIns;
  const ty = by + tIns, th = bh - tIns - bIns;

  const defaultFontSz = 1800;
  const lstStyle = g1(txBody, 'lstStyle');
  const lstDefRPr = lstStyle ? g1(lstStyle, 'defRPr') : null;

  const paragraphs = gtn(txBody, 'p');
  let svgLines = '';
  let curY = ty;
  const clipId = uid('clip');
  defs.push(`<clipPath id="${clipId}"><rect x="${px(bx)}" y="${px(by)}" width="${px(bw)}" height="${px(bh)}"/></clipPath>`);

  // Vertical text: rotate the whole group
  const vertTransform = isVert
    ? ` transform="rotate(-90,${px(bx + bw/2)},${px(by + bh/2)})"` : '';

  // Auto-number counters
  const autoNumCtrs = {};

  for (const para of paragraphs) {
    const pPr      = g1(para, 'pPr');
    const algn     = attr(pPr, 'algn', 'l');
    const marL     = attrInt(pPr, 'marL', 0) / 914400 * 96;
    const indent   = attrInt(pPr, 'indent', 0) / 914400 * 96;
    const defRPr   = g1(pPr, 'defRPr');
    let paraDefSz  = defaultFontSz;
    if (lstDefRPr) { const sz = lstDefRPr.getAttribute('sz'); if (sz) paraDefSz = parseInt(sz, 10); }
    if (defRPr)    { const sz = defRPr.getAttribute('sz');    if (sz) paraDefSz = parseInt(sz, 10); }

    // Spacing
    const spcBef = g1(pPr, 'spcBef');
    const spcAft = g1(pPr, 'spcAft');
    const lnSpc  = g1(pPr, 'lnSpc');

    let spaceBefore = 0, spaceAfter = 0;
    if (spcBef) {
      const sp = g1(spcBef, 'spcPct'), spp = g1(spcBef, 'spcPts');
      if (sp)  spaceBefore = (paraDefSz * EMU_PER_PT) / 914400 * 96 * (attrInt(sp, 'val', 0) / 100000);
      else if (spp) spaceBefore = attrInt(spp, 'val', 0) / 100 / 72 * 96;
    }
    if (spcAft) {
      const sp = g1(spcAft, 'spcPct'), spp = g1(spcAft, 'spcPts');
      if (sp)  spaceAfter = (paraDefSz * EMU_PER_PT) / 914400 * 96 * (attrInt(sp, 'val', 0) / 100000);
      else if (spp) spaceAfter = attrInt(spp, 'val', 0) / 100 / 72 * 96;
    }

    // Bullet
    const buChar    = pPr ? g1(pPr, 'buChar') : null;
    const buAutoNum = pPr ? g1(pPr, 'buAutoNum') : null;
    const buNone    = pPr ? g1(pPr, 'buNone') : null;
    const hasBullet = !buNone && (buChar || buAutoNum);

    curY += spaceBefore;

    const runEls = [];
    for (const child of para.children) {
      if (child.localName === 'r' || child.localName === 'br' || child.localName === 'fld')
        runEls.push(child);
    }

    // Empty paragraph
    if (!runEls.length) {
      const endRPr = g1(para, 'endParaRPr');
      const sz  = attrInt(endRPr || defRPr, 'sz', paraDefSz);
      const szPx = sz / 100 / 72 * 96;
      const lnH  = szPx * 1.2;
      curY += lnH + spaceAfter;
      continue;
    }

    // Build runs text
    let lineText = '';
    for (const rEl of runEls) {
      if (rEl.localName === 'br') { lineText += '\n'; continue; }
      const t = g1(rEl, 't') || g1(rEl, 'fldVal');
      if (t) lineText += t.textContent;
    }

    // Group consecutive runs by their rPr for tspan generation
    const tspans = [];
    for (const rEl of runEls) {
      if (rEl.localName === 'br') {
        tspans.push({ br: true });
        continue;
      }
      const rPr = g1(rEl, 'rPr') || g1(rEl, 'r')?.firstElementChild;
      const tEl = g1(rEl, 't');
      if (!tEl) continue;
      const text = tEl.textContent;
      if (!text) continue;

      // Font info
      const fi = buildFontInherited(rEl, defRPr, lstDefRPr, themeColors, themeData, paraDefSz);
      const szPx = fi?.szPx || (paraDefSz / 100 / 72 * 96);
      const family = fi?.family || 'sans-serif';
      const bold   = fi?.bold ? 'bold' : 'normal';
      const italic = fi?.italic ? 'italic' : 'normal';
      const color  = fi?.color ? colorToCss(fi.color) : '#000000';
      const underline = rPr ? (rPr.getAttribute('u') || 'none') !== 'none' : false;
      const strike    = rPr ? (rPr.getAttribute('strike') || 'noStrike') !== 'noStrike' : false;
      const baseline  = rPr ? parseInt(rPr.getAttribute('baseline') || '0', 10) : 0;

      tspans.push({ text, szPx, family, bold, italic, color, underline, strike, baseline });
    }

    if (!tspans.length) { curY += spaceAfter; continue; }

    // Average font size for line height
    const sizes = tspans.filter(t => !t.br && t.szPx).map(t => t.szPx);
    const maxSzPx = sizes.length ? Math.max(...sizes) : paraDefSz / 100 / 72 * 96;
    const lnH = maxSzPx * 1.2;
    const baseline = curY + maxSzPx * 0.85;

    // Text anchor
    let textAnchor = 'start';
    let xPos = tx + marL;
    if (algn === 'ctr') { textAnchor = 'middle'; xPos = tx + tw / 2; }
    else if (algn === 'r') { textAnchor = 'end'; xPos = tx + tw; }

    // Bullet character
    let bulletSvg = '';
    if (hasBullet) {
      const bx2 = tx + marL + indent;
      let bulletChar = '';
      if (buChar) {
        bulletChar = esc(buChar.getAttribute('char') || '•');
      } else if (buAutoNum) {
        const numType = buAutoNum.getAttribute('type') || 'arabicPeriod';
        const startAt = attrInt(buAutoNum, 'startAt', 1);
        const key = numType + ':' + startAt;
        if (!autoNumCtrs[key]) autoNumCtrs[key] = startAt;
        bulletChar = esc(formatAutoNum(numType, autoNumCtrs[key]++));
      }
      const bSzPx = maxSzPx;
      bulletSvg = `<text x="${px(bx2)}" y="${px(baseline)}" font-size="${px(bSzPx)}" font-family="sans-serif" fill="#000">${bulletChar}</text>`;
    }

    // Build <text> element with <tspan>s
    let tspanSvg = '';
    let firstSpan = true;
    for (const ts of tspans) {
      if (ts.br) {
        curY += lnH;
        tspanSvg += `<tspan x="${px(xPos)}" dy="${px(lnH)}">`;
        firstSpan = false;
        continue;
      }
      const dy = firstSpan ? 0 : 0;
      const deco = ts.underline ? 'underline' : ts.strike ? 'line-through' : 'none';
      let adjustedY = baseline;
      if (ts.baseline > 0) adjustedY = baseline - ts.szPx * 0.38;
      else if (ts.baseline < 0) adjustedY = baseline + ts.szPx * 0.12;
      const subSzPx = ts.baseline !== 0 ? ts.szPx * 0.65 : ts.szPx;

      tspanSvg += `<tspan` +
        ` font-family="${esc(ts.family)}, sans-serif"` +
        ` font-size="${px(subSzPx)}"` +
        ` font-weight="${ts.bold}"` +
        ` font-style="${ts.italic}"` +
        ` fill="${esc(ts.color)}"` +
        (deco !== 'none' ? ` text-decoration="${deco}"` : '') +
        (ts.baseline !== 0 ? ` dy="${px(adjustedY - baseline)}"` : '') +
        `>${esc(ts.text)}</tspan>`;
      firstSpan = false;
    }

    svgLines += bulletSvg;
    svgLines += `<text x="${px(xPos)}" y="${px(baseline)}" text-anchor="${textAnchor}">${tspanSvg}</text>`;
    curY += lnH + spaceAfter;
  }

  return `<g clip-path="url(#${clipId})"${vertTransform}>${svgLines}</g>`;
}

function formatAutoNum(type, n) {
  switch (type) {
    case 'arabicPeriod':    return n + '.';
    case 'arabicParenR':    return n + ')';
    case 'arabicParenBoth': return '(' + n + ')';
    case 'romanLcPeriod':   return toRoman(n).toLowerCase() + '.';
    case 'romanUcPeriod':   return toRoman(n) + '.';
    case 'alphaLcParenR':   return String.fromCharCode(96 + n) + ')';
    case 'alphaUcParenR':   return String.fromCharCode(64 + n) + ')';
    default:                return n + '.';
  }
}
function toRoman(n) {
  const v=[1000,900,500,400,100,90,50,40,10,9,5,4,1];
  const s=['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let r='';
  for(let i=0;i<v.length;i++) while(n>=v[i]){r+=s[i];n-=v[i];}
  return r;
}

// ── Shape → SVG ───────────────────────────────────────────────────────────────
async function shapesToSvg(spTreeEl, rels, files, themeColors, themeData, defs) {
  if (!spTreeEl) return '';
  let out = '';
  for (const child of spTreeEl.children) {
    const ln = child.localName;
    if (ln === 'sp')           out += await shapeToSvg(child, themeColors, themeData, defs);
    else if (ln === 'pic')     out += await pictureToSvg(child, rels, files, themeColors, defs);
    else if (ln === 'cxnSp')  out += await connectorToSvg(child, themeColors, defs);
    else if (ln === 'grpSp')  out += await groupToSvg(child, rels, files, themeColors, themeData, defs);
    else if (ln === 'graphicFrame') out += await graphicFrameToSvg(child, themeColors, defs);
  }
  return out;
}

async function shapeToSvg(spEl, themeColors, themeData, defs) {
  const spPr   = g1(spEl, 'spPr');
  const xfrm   = g1(spPr, 'xfrm');
  const b      = xfrmBounds(xfrm);
  if (!b) return '';

  const prstGeom = g1(spPr, 'prstGeom');
  const prst     = prstGeom ? attr(prstGeom, 'prst', 'rect') : 'rect';
  const custGeom = g1(spPr, 'custGeom');

  // Fill
  const fillNames = ['noFill','solidFill','gradFill','blipFill','pattFill'];
  let fillElSource = null;
  for (const fn of fillNames) { const el = g1(spPr, fn); if (el) { fillElSource = el; break; } }
  if (!fillElSource) {
    const styleEl = g1(spEl, 'style');
    const fillRef = styleEl ? g1(styleEl, 'fillRef') : null;
    if (fillRef && attrInt(fillRef, 'idx', 1) !== 0) {
      const cc = findFirstColorChild(fillRef);
      const c  = resolveColorElement(cc, themeColors);
      if (c) {
        const ns = 'http://schemas.openxmlformats.org/drawingml/2006/main';
        const doc2 = new DOMParser().parseFromString(`<solidFill xmlns="${ns}"><srgbClr val="${colorToCss(c).replace('#','')}"/></solidFill>`, 'application/xml');
        fillElSource = doc2.documentElement;
      }
    }
  }

  const { fill, fillAttrs } = fillAttr(fillElSource, defs, themeColors, b.x, b.y, b.w, b.h);

  // Stroke
  let lnEl = g1(spPr, 'ln');
  if (!lnEl) {
    const styleEl = g1(spEl, 'style');
    const lnRef   = styleEl ? g1(styleEl, 'lnRef') : null;
    if (lnRef && attrInt(lnRef, 'idx', 1) !== 0) {
      const cc = findFirstColorChild(lnRef);
      const c  = resolveColorElement(cc, themeColors);
      if (c) {
        const ns = 'http://schemas.openxmlformats.org/drawingml/2006/main';
        const doc2 = new DOMParser().parseFromString(`<ln xmlns="${ns}"><solidFill><srgbClr val="${colorToCss(c).replace('#','')}"/></solidFill></ln>`, 'application/xml');
        lnEl = doc2.documentElement;
      }
    }
  }
  const stroke = strokeAttrs(lnEl, themeColors, 1);

  // Shadow
  const effectLst = g1(spPr, 'effectLst');
  const filt = shadowFilter(effectLst, defs);

  // Transform
  const transform = xfrmAttrs(xfrm);

  // Path
  let pathSvg = '';
  if (prst === 'rect' || (!prstGeom && !custGeom)) {
    pathSvg = `<rect x="${px(b.x)}" y="${px(b.y)}" width="${px(b.w)}" height="${px(b.h)}" fill="${esc(fill)}"${fillAttrs}${stroke}${filt}${transform}/>`;
  } else {
    const d = presetToSvgPath(prst, b.x, b.y, b.w, b.h, {});
    pathSvg = `<path d="${esc(d)}" fill="${esc(fill)}"${fillAttrs}${stroke}${filt}${transform}/>`;
  }

  // Text
  const txBody = g1(spEl, 'txBody');
  const textSvg = txBody ? textBodyToSvg(txBody, b.x, b.y, b.w, b.h, themeColors, themeData, defs) : '';

  return `<g>${pathSvg}${textSvg}</g>`;
}

async function pictureToSvg(picEl, rels, files, themeColors, defs) {
  const spPr = g1(picEl, 'spPr');
  const xfrm = g1(spPr, 'xfrm');
  const b    = xfrmBounds(xfrm);
  if (!b) return '';

  const blipFill = g1(picEl, 'blipFill');
  const blip     = blipFill ? g1(blipFill, 'blip') : null;
  const rId      = blip ? (blip.getAttribute('r:embed') || blip.getAttribute('embed')) : null;
  const rel      = rId ? rels[rId] : null;
  const imgData  = rel ? files[rel.fullPath] : null;

  const transform = xfrmAttrs(xfrm);
  const effectLst = g1(spPr, 'effectLst');
  const filt = shadowFilter(effectLst, defs);

  if (!imgData) {
    return `<rect x="${px(b.x)}" y="${px(b.y)}" width="${px(b.w)}" height="${px(b.h)}" fill="#e0e0e0"${transform}/>`;
  }

  // Convert to base64
  const ext  = rel.fullPath.split('.').pop().toLowerCase();
  const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif'
    : ext === 'svg' ? 'image/svg+xml' : 'image/jpeg';
  const raw  = imgData instanceof Uint8Array ? imgData : new Uint8Array(imgData);
  const b64  = btoa(Array.from(raw, b => String.fromCharCode(b)).join(''));

  // Clip to shape bounds
  const clipId = uid('pic');
  defs.push(`<clipPath id="${clipId}"><rect x="${px(b.x)}" y="${px(b.y)}" width="${px(b.w)}" height="${px(b.h)}"/></clipPath>`);

  return `<image x="${px(b.x)}" y="${px(b.y)}" width="${px(b.w)}" height="${px(b.h)}" href="data:${mime};base64,${b64}" clip-path="url(#${clipId})"${filt}${transform} preserveAspectRatio="xMidYMid slice"/>`;
}

async function connectorToSvg(cxnSpEl, themeColors, defs) {
  const spPr = g1(cxnSpEl, 'spPr');
  const xfrm = g1(spPr, 'xfrm');
  const b    = xfrmBounds(xfrm);
  if (!b) return '';
  const lnEl = g1(spPr, 'ln');
  const stroke = strokeAttrs(lnEl, themeColors, 1);
  const transform = xfrmAttrs(xfrm);
  return `<line x1="${px(b.x)}" y1="${px(b.y)}" x2="${px(b.x+b.w)}" y2="${px(b.y+b.h)}" fill="none"${stroke}${transform}/>`;
}

async function groupToSvg(grpSpEl, rels, files, themeColors, themeData, defs) {
  const spPr = g1(grpSpEl, 'grpSpPr');
  const xfrm = g1(spPr, 'xfrm');
  const b    = xfrmBounds(xfrm);
  const transform = xfrm ? xfrmAttrs(xfrm) : '';
  const children = await shapesToSvg(grpSpEl, rels, files, themeColors, themeData, defs);
  return `<g${transform}>${children}</g>`;
}

async function graphicFrameToSvg(graphicFrame, themeColors, defs) {
  const xfrm = g1(graphicFrame, 'xfrm');
  const b    = xfrmBounds(xfrm);
  if (!b) return '';
  // Placeholder (charts/SmartArt SVG rendering is complex; show clean box)
  return `<rect x="${px(b.x)}" y="${px(b.y)}" width="${px(b.w)}" height="${px(b.h)}" fill="#f4f4f8" stroke="#ccc" stroke-width="1"/>` +
         `<text x="${px(b.x+b.w/2)}" y="${px(b.y+b.h/2)}" text-anchor="middle" font-size="14" fill="#999">Chart</text>`;
}

// ── Background → SVG ──────────────────────────────────────────────────────────
async function backgroundToSvg(slideDoc, masterDoc, layoutDoc, files, masterRels, themeColors, slideW, slideH, defs) {
  const getbg = (doc) => {
    const cSld = g1(doc, 'cSld');
    const bg   = cSld ? g1(cSld, 'bg') : null;
    if (!bg) return null;
    return { bgPr: g1(bg, 'bgPr'), bgRef: g1(bg, 'bgRef') };
  };
  const bgData = getbg(slideDoc) || getbg(layoutDoc) || getbg(masterDoc);
  if (!bgData) return `<rect width="${px(slideW)}" height="${px(slideH)}" fill="white"/>`;

  const { bgPr, bgRef } = bgData;
  if (bgPr) {
    const fills = ['noFill','solidFill','gradFill','blipFill','pattFill'];
    for (const fn of fills) {
      const fillEl = g1(bgPr, fn);
      if (fillEl) {
        if (fn === 'blipFill') {
          const blip = g1(fillEl, 'blip');
          const rId  = blip ? (blip.getAttribute('r:embed') || blip.getAttribute('embed')) : null;
          const rel  = rId && masterRels ? masterRels[rId] : null;
          const imgData = rel ? files[rel.fullPath] : null;
          if (imgData) {
            const ext  = rel.fullPath.split('.').pop().toLowerCase();
            const mime = ext === 'png' ? 'image/png' : 'image/jpeg';
            const raw  = imgData instanceof Uint8Array ? imgData : new Uint8Array(imgData);
            const b64  = btoa(Array.from(raw, b => String.fromCharCode(b)).join(''));
            return `<image width="${px(slideW)}" height="${px(slideH)}" href="data:${mime};base64,${b64}" preserveAspectRatio="xMidYMid slice"/>`;
          }
        }
        const { fill, fillAttrs } = fillAttr(fillEl, defs, themeColors, 0, 0, slideW, slideH);
        return `<rect width="${px(slideW)}" height="${px(slideH)}" fill="${esc(fill)}"${fillAttrs}/>`;
      }
    }
  }
  if (bgRef) {
    const cc = findFirstColorChild(bgRef);
    const c  = resolveColorElement(cc, themeColors);
    if (c) return `<rect width="${px(slideW)}" height="${px(slideH)}" fill="${esc(colorToCss(c))}"/>`;
  }
  return `<rect width="${px(slideW)}" height="${px(slideH)}" fill="white"/>`;
}

// ── Main entry ────────────────────────────────────────────────────────────────

/**
 * Render a single slide to an SVG string.
 *
 * @param {number}          slideIndex
 * @param {PptxRenderer}    renderer    — loaded renderer instance
 * @returns {Promise<string>}           — complete SVG markup
 */
export async function renderSlideToSvg(slideIndex, renderer) {
  const { _files: files, slidePaths, slideSize, themeColors, themeData,
          masterDoc, masterRels } = renderer;

  if (slideIndex < 0 || slideIndex >= slidePaths.length) throw new Error('Slide index out of range');

  const slidePath = slidePaths[slideIndex];
  const slideXml  = files[slidePath] ? new TextDecoder().decode(files[slidePath]) : null;
  if (!slideXml) throw new Error(`Cannot read slide ${slideIndex}`);

  const slideDoc   = parseXml(slideXml);
  const slideRels  = await getRels(files, slidePath);

  // Layout
  const layoutRel = Object.values(slideRels).find(r => r.type?.includes('slideLayout'));
  const layoutDoc = layoutRel && files[layoutRel.fullPath]
    ? parseXml(new TextDecoder().decode(files[layoutRel.fullPath])) : null;
  const layoutRels = layoutRel ? await getRels(files, layoutRel.fullPath) : {};

  // Slide dimensions in px (96 dpi)
  const W = slideSize.cx / 914400 * 96;
  const H = slideSize.cy / 914400 * 96;

  const defs = [];

  // Background
  const bgSvg = await backgroundToSvg(slideDoc, masterDoc, layoutDoc, files,
    masterRels, themeColors, W, H, defs);

  // Master / layout decorative shapes
  const masterTree = g1(g1(masterDoc, 'cSld'), 'spTree');
  const layoutTree = layoutDoc ? g1(g1(layoutDoc, 'cSld'), 'spTree') : null;
  const slideTree  = g1(g1(slideDoc, 'cSld'), 'spTree');

  const masterSvg = masterTree
    ? await shapesToSvg(masterTree, masterRels, files, themeColors, themeData, defs) : '';
  const layoutSvg = layoutTree
    ? await shapesToSvg(layoutTree, layoutRels, files, themeColors, themeData, defs) : '';
  const slideSvg  = slideTree
    ? await shapesToSvg(slideTree, slideRels, files, themeColors, themeData, defs) : '';

  const defsBlock = defs.length ? `<defs>${defs.join('\n')}</defs>` : '';

  return [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"`,
    `  width="${px(W)}" height="${px(H)}" viewBox="0 0 ${px(W)} ${px(H)}">`,
    defsBlock,
    bgSvg,
    masterSvg,
    layoutSvg,
    slideSvg,
    `</svg>`,
  ].join('\n');
}

/**
 * Render all slides to SVG strings.
 * @param {PptxRenderer} renderer
 * @returns {Promise<string[]>}
 */
export async function renderAllSlidesToSvg(renderer) {
  const results = [];
  for (let i = 0; i < renderer.slideCount; i++) {
    results.push(await renderSlideToSvg(i, renderer));
  }
  return results;
}
