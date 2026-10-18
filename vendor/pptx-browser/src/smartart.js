/**
 * smartart.js — SmartArt / Diagram renderer.
 *
 * Parses ppt/diagrams/data*.xml and layout*.xml and renders the 12 most
 * common SmartArt layout families. Unknown layouts render a clean node-list
 * fallback that shows all the text content.
 *
 * Supported layout families (detected from layout XML typ attribute):
 *   process / chevronList / accentProcess / blockList → Process flow
 *   cycle / continuousCycle / blockCycle              → Cycle
 *   hierarchy / orgChart / horizontalOrg             → Hierarchy / org chart
 *   radial / accentedRadial / cycleMatrix             → Radial
 *   pyramid / invertedPyramid                        → Pyramid
 *   funnel                                           → Funnel
 *   venn / linearVenn                                → Venn
 *   list / verticalBoxList / pictureAccentList        → List
 *   matrix / accentedMatrix                          → 2×2 matrix
 *   relationship / divergingRadial                   → Relationship
 *   stapledDocument / squareAccentList               → fallback list
 */

import { g1, gtn, attr } from './utils.js';

// ── Colour palette ───────────────────────────────────────────────────────────
const PALETTE = [
  '#4472C4', '#ED7D31', '#A9D18E', '#FF0000',
  '#FFC000', '#5B9BD5', '#70AD47', '#C00000',
  '#7030A0', '#00B0F0',
];

function nodeColor(i, themeColors) {
  const key = `accent${(i % 6) + 1}`;
  if (themeColors[key]) return '#' + themeColors[key];
  return PALETTE[i % PALETTE.length];
}

function lighten(hex, amount = 0.45) {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const lr = Math.round(r + (255 - r) * amount);
  const lg = Math.round(g + (255 - g) * amount);
  const lb = Math.round(b + (255 - b) * amount);
  return `rgb(${lr},${lg},${lb})`;
}

// ── XML helpers ──────────────────────────────────────────────────────────────

/** Extract all text from a SmartArt node element (pt). */
function nodeText(ptEl) {
  const texts = [];
  for (const t of gtn(ptEl, 't')) {
    const txt = t.textContent.trim();
    if (txt) texts.push(txt);
  }
  return texts.join(' ');
}

/** Read all data nodes (non-connector pts) from diagram data XML. */
function readNodes(dataDoc) {
  if (!dataDoc) return [];
  const ptLst = g1(dataDoc, 'ptLst');
  if (!ptLst) return [];
  const nodes = [];
  for (const pt of gtn(ptLst, 'pt')) {
    const type = attr(pt, 'type', 'node');
    if (type === 'parTrans' || type === 'sibTrans') continue;
    const text = nodeText(pt);
    if (text || type === 'node') {
      nodes.push({ id: attr(pt, 'modelId', ''), type, text });
    }
  }
  return nodes.filter(n => n.text);
}

/** Detect layout type from layout XML. */
function detectLayout(layoutDoc) {
  if (!layoutDoc) return 'list';
  const diagDef = g1(layoutDoc, 'layoutDef') || layoutDoc;
  const typ = attr(diagDef, 'uniqueId', '')
    || attr(diagDef, 'defStyle', '')
    || '';
  const t = typ.toLowerCase();
  if (t.includes('chevron') || t.includes('arrowprocess') || t.includes('process'))
    return 'process';
  if (t.includes('cycle') || t.includes('continuouscycle'))
    return 'cycle';
  if (t.includes('hierarchy') || t.includes('orgchart') || t.includes('org'))
    return 'hierarchy';
  if (t.includes('radial') || t.includes('diverging'))
    return 'radial';
  if (t.includes('pyramid') || t.includes('invertedpyramid'))
    return 'pyramid';
  if (t.includes('funnel'))
    return 'funnel';
  if (t.includes('venn'))
    return 'venn';
  if (t.includes('matrix'))
    return 'matrix';
  if (t.includes('list') || t.includes('bullet'))
    return 'list';
  if (t.includes('relationship') || t.includes('balance'))
    return 'relationship';
  return 'list';
}

// ── Text rendering ───────────────────────────────────────────────────────────

function drawText(ctx, text, x, y, maxW, maxH, size, color = '#fff', align = 'center') {
  if (!text) return;
  ctx.save();
  ctx.fillStyle = color;
  ctx.textAlign = align;
  ctx.textBaseline = 'middle';
  const words = text.split(' ');
  const lineH = size * 1.3;
  let lines = [];
  let line = '';
  for (const w of words) {
    const test = line ? line + ' ' + w : w;
    if (ctx.measureText(test).width > maxW - 4 && line) {
      lines.push(line);
      line = w;
    } else {
      line = test;
    }
  }
  if (line) lines.push(line);
  // Limit lines to fit height
  const maxLines = Math.max(1, Math.floor(maxH / lineH));
  if (lines.length > maxLines) lines = lines.slice(0, maxLines);
  const startY = y - (lines.length - 1) * lineH / 2;
  for (let i = 0; i < lines.length; i++) {
    ctx.fillText(lines[i], x, startY + i * lineH);
  }
  ctx.restore();
}

function autoFontSize(text, w, h, maxSz = 14) {
  const approxCharsPerLine = Math.floor(w / (maxSz * 0.55));
  const words = text.split(' ');
  const lines = Math.ceil(words.join(' ').length / Math.max(approxCharsPerLine, 1));
  const byH = h / (lines * 1.4);
  return Math.max(8, Math.min(maxSz, byH));
}

// ── Shared shape primitives ──────────────────────────────────────────────────

function fillRoundRect(ctx, x, y, w, h, r, fill, stroke) {
  ctx.save();
  ctx.beginPath();
  r = Math.min(r, w / 2, h / 2);
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
  if (fill) { ctx.fillStyle = fill; ctx.fill(); }
  if (stroke) { ctx.strokeStyle = stroke; ctx.lineWidth = 1; ctx.stroke(); }
  ctx.restore();
}

function drawArrow(ctx, x1, y1, x2, y2, color, lw = 1.5) {
  const angle = Math.atan2(y2 - y1, x2 - x1);
  const headLen = 8;
  ctx.save();
  ctx.strokeStyle = color;
  ctx.fillStyle = color;
  ctx.lineWidth = lw;
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(x2, y2);
  ctx.lineTo(x2 - headLen * Math.cos(angle - 0.4), y2 - headLen * Math.sin(angle - 0.4));
  ctx.lineTo(x2 - headLen * Math.cos(angle + 0.4), y2 - headLen * Math.sin(angle + 0.4));
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

// ── Layout renderers ─────────────────────────────────────────────────────────

function renderProcess(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (!nodes.length) return;
  const N = nodes.length;
  const itemW = w / N;
  const pad = Math.min(itemW * 0.08, 10 * scale);
  const arrowW = Math.min(itemW * 0.12, 18 * scale);
  const boxW = itemW - arrowW - pad;
  const boxH = h * 0.7;
  const boxY = y + (h - boxH) / 2;

  for (let i = 0; i < N; i++) {
    const color = nodeColor(i, themeColors);
    const bx = x + itemW * i + pad / 2;
    const isLast = i === N - 1;

    // Chevron arrow shape
    ctx.save();
    ctx.beginPath();
    const tipX = bx + boxW + (isLast ? 0 : arrowW);
    ctx.moveTo(bx, boxY);
    ctx.lineTo(bx + boxW, boxY);
    if (!isLast) ctx.lineTo(tipX, boxY + boxH / 2);
    ctx.lineTo(bx + boxW, boxY + boxH);
    ctx.lineTo(bx, boxY + boxH);
    if (i > 0) ctx.lineTo(bx + arrowW, boxY + boxH / 2);
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();
    ctx.restore();

    const fs = autoFontSize(nodes[i].text, boxW - arrowW, boxH) * scale;
    ctx.font = `${fs}px sans-serif`;
    const cx2 = bx + (boxW + (i > 0 ? arrowW : 0)) / 2;
    drawText(ctx, nodes[i].text, cx2, boxY + boxH / 2, boxW, boxH, fs);
  }
}

function renderCycle(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (!nodes.length) return;
  const N = nodes.length;
  const pcx = x + w / 2;
  const pcy = y + h / 2;
  const orbitR = Math.min(w, h) * 0.33;
  const nodeR  = Math.min(orbitR * 0.32, 40 * scale);

  // Draw connecting ring
  ctx.save();
  ctx.strokeStyle = '#ccc';
  ctx.lineWidth = 2 * scale;
  ctx.setLineDash([4 * scale, 4 * scale]);
  ctx.beginPath();
  ctx.arc(pcx, pcy, orbitR, 0, Math.PI * 2);
  ctx.stroke();
  ctx.setLineDash([]);
  ctx.restore();

  for (let i = 0; i < N; i++) {
    const angle = (i / N) * Math.PI * 2 - Math.PI / 2;
    const nx = pcx + orbitR * Math.cos(angle);
    const ny = pcy + orbitR * Math.sin(angle);
    const color = nodeColor(i, themeColors);

    // Circle node
    ctx.save();
    ctx.beginPath();
    ctx.arc(nx, ny, nodeR, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.fill();
    ctx.restore();

    // Curved arrow to next node
    if (N > 1) {
      const nextAngle = ((i + 1) / N) * Math.PI * 2 - Math.PI / 2;
      const midAngle = (angle + nextAngle) / 2;
      const ax1 = pcx + orbitR * Math.cos(angle + 0.15);
      const ay1 = pcy + orbitR * Math.sin(angle + 0.15);
      const ax2 = pcx + orbitR * Math.cos(nextAngle - 0.15);
      const ay2 = pcy + orbitR * Math.sin(nextAngle - 0.15);
      ctx.save();
      ctx.strokeStyle = color + '99';
      ctx.lineWidth = 1.5 * scale;
      ctx.beginPath();
      ctx.moveTo(ax1, ay1);
      ctx.quadraticCurveTo(pcx + orbitR * 1.1 * Math.cos(midAngle),
                           pcy + orbitR * 1.1 * Math.sin(midAngle), ax2, ay2);
      ctx.stroke();
      ctx.restore();
    }

    const fs = autoFontSize(nodes[i].text, nodeR * 1.7, nodeR * 1.2) * scale;
    ctx.font = `${fs}px sans-serif`;
    drawText(ctx, nodes[i].text, nx, ny, nodeR * 2, nodeR * 1.8, fs);
  }
}

function renderHierarchy(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (!nodes.length) return;
  const root = nodes[0];
  const children = nodes.slice(1);
  const N = Math.max(children.length, 1);

  const rowH = h / (children.length > 0 ? 2.8 : 1.2);
  const boxH = rowH * 0.75;
  const rootW = Math.min(w * 0.3, 160 * scale);
  const rootH = boxH;
  const rootX = x + (w - rootW) / 2;
  const rootY = y + (rowH - rootH) / 2;
  const color0 = nodeColor(0, themeColors);

  fillRoundRect(ctx, rootX, rootY, rootW, rootH, 6 * scale, color0, null);
  const fs0 = autoFontSize(root.text, rootW - 10 * scale, rootH) * scale;
  ctx.font = `bold ${fs0}px sans-serif`;
  drawText(ctx, root.text, rootX + rootW / 2, rootY + rootH / 2, rootW - 10, rootH, fs0);

  if (!children.length) return;

  const childW = (w - 20 * scale) / N;
  const childRowY = y + rowH * 1.5;

  // Connector from root
  const lineStartY = rootY + rootH;
  const lineEndY   = childRowY - 4 * scale;
  ctx.save();
  ctx.strokeStyle = '#999';
  ctx.lineWidth = 1.5 * scale;
  ctx.beginPath();
  ctx.moveTo(rootX + rootW / 2, lineStartY);
  ctx.lineTo(rootX + rootW / 2, (lineStartY + lineEndY) / 2);
  ctx.lineTo(x + 10 * scale, (lineStartY + lineEndY) / 2);
  ctx.lineTo(x + w - 10 * scale, (lineStartY + lineEndY) / 2);
  ctx.stroke();
  ctx.restore();

  for (let i = 0; i < children.length; i++) {
    const cx2 = x + childW * i + 8 * scale;
    const cy2 = childRowY;
    const cw2 = childW - 16 * scale;
    const ch2 = boxH;
    const color = nodeColor(i + 1, themeColors);

    // Vertical line from horizontal bus
    ctx.save();
    ctx.strokeStyle = '#999';
    ctx.lineWidth = 1.5 * scale;
    ctx.beginPath();
    ctx.moveTo(cx2 + cw2 / 2, (lineStartY + lineEndY) / 2);
    ctx.lineTo(cx2 + cw2 / 2, cy2);
    ctx.stroke();
    ctx.restore();

    fillRoundRect(ctx, cx2, cy2, cw2, ch2, 5 * scale, color, null);
    const fs = autoFontSize(children[i].text, cw2 - 8 * scale, ch2) * scale;
    ctx.font = `${fs}px sans-serif`;
    drawText(ctx, children[i].text, cx2 + cw2 / 2, cy2 + ch2 / 2, cw2 - 8, ch2, fs);
  }
}

function renderRadial(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (!nodes.length) return;
  const center = nodes[0];
  const spokes = nodes.slice(1);
  const pcx = x + w / 2;
  const pcy = y + h / 2;
  const coreR = Math.min(w, h) * 0.18;
  const orbitR = Math.min(w, h) * 0.36;
  const nodeR  = Math.min(orbitR * 0.28, 38 * scale);
  const N = spokes.length || 1;

  // Center circle
  const c0 = nodeColor(0, themeColors);
  ctx.save();
  ctx.beginPath();
  ctx.arc(pcx, pcy, coreR, 0, Math.PI * 2);
  ctx.fillStyle = c0;
  ctx.fill();
  ctx.restore();
  const cfs = autoFontSize(center.text, coreR * 1.6, coreR * 1.2) * scale;
  ctx.font = `bold ${cfs}px sans-serif`;
  drawText(ctx, center.text, pcx, pcy, coreR * 2, coreR * 2, cfs);

  for (let i = 0; i < N; i++) {
    const angle = (i / N) * Math.PI * 2 - Math.PI / 2;
    const nx = pcx + orbitR * Math.cos(angle);
    const ny = pcy + orbitR * Math.sin(angle);
    const color = nodeColor(i + 1, themeColors);

    // Spoke line
    const lx1 = pcx + coreR * Math.cos(angle);
    const ly1 = pcy + coreR * Math.sin(angle);
    const lx2 = nx - nodeR * Math.cos(angle);
    const ly2 = ny - nodeR * Math.sin(angle);
    drawArrow(ctx, lx1, ly1, lx2, ly2, color + '99', 1.5 * scale);

    ctx.save();
    ctx.beginPath();
    ctx.arc(nx, ny, nodeR, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.fill();
    ctx.restore();

    const fs = autoFontSize(spokes[i].text, nodeR * 1.7, nodeR * 1.2) * scale;
    ctx.font = `${fs}px sans-serif`;
    drawText(ctx, spokes[i].text, nx, ny, nodeR * 2, nodeR * 1.8, fs);
  }
}

function renderPyramid(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (!nodes.length) return;
  const N = nodes.length;
  const layerH = h / N;
  for (let i = 0; i < N; i++) {
    const t = (N - i) / N;         // fraction of base width at this layer
    const layerW = w * t;
    const lx = x + (w - layerW) / 2;
    const ly = y + i * layerH;
    const color = nodeColor(i, themeColors);

    ctx.save();
    ctx.beginPath();
    ctx.moveTo(lx, ly + layerH);
    ctx.lineTo(lx + layerW, ly + layerH);
    const topW = w * ((N - i - 1) / N);
    ctx.lineTo(x + (w + topW) / 2, ly);
    ctx.lineTo(x + (w - topW) / 2, ly);
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 2 * scale;
    ctx.fill();
    ctx.stroke();
    ctx.restore();

    const fs = autoFontSize(nodes[i].text, layerW * 0.6, layerH * 0.7) * scale;
    ctx.font = `${fs}px sans-serif`;
    drawText(ctx, nodes[i].text, x + w / 2, ly + layerH / 2, layerW * 0.6, layerH * 0.7, fs);
  }
}

function renderFunnel(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (!nodes.length) return;
  const N = nodes.length;
  const layerH = h / N;

  for (let i = 0; i < N; i++) {
    // Funnel: wide at top, narrow at bottom
    const topFrac = 1 - (i / N) * 0.55;
    const botFrac = 1 - ((i + 1) / N) * 0.55;
    const topW = w * topFrac;
    const botW = w * botFrac;
    const lx1 = x + (w - topW) / 2;
    const lx2 = x + (w - botW) / 2;
    const ly = y + i * layerH;
    const color = nodeColor(i, themeColors);

    ctx.save();
    ctx.beginPath();
    ctx.moveTo(lx1, ly);
    ctx.lineTo(lx1 + topW, ly);
    ctx.lineTo(lx2 + botW, ly + layerH);
    ctx.lineTo(lx2, ly + layerH);
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 2 * scale;
    ctx.fill();
    ctx.stroke();
    ctx.restore();

    const mw = (topW + botW) / 2;
    const fs = autoFontSize(nodes[i].text, mw * 0.8, layerH * 0.7) * scale;
    ctx.font = `${fs}px sans-serif`;
    drawText(ctx, nodes[i].text, x + w / 2, ly + layerH / 2, mw * 0.8, layerH * 0.7, fs);
  }
}

function renderVenn(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (!nodes.length) return;
  const N = Math.min(nodes.length, 4); // Venn beyond 4 is unreadable
  const pcx = x + w / 2;
  const pcy = y + h / 2;
  const cr  = Math.min(w, h) * (N <= 2 ? 0.32 : 0.28);
  const spread = cr * 0.65;

  const angles = [];
  for (let i = 0; i < N; i++) angles.push((i / N) * Math.PI * 2 - Math.PI / 2);

  for (let i = 0; i < N; i++) {
    const nx = N > 1 ? pcx + spread * Math.cos(angles[i]) : pcx;
    const ny = N > 1 ? pcy + spread * Math.sin(angles[i]) : pcy;
    const color = nodeColor(i, themeColors);

    ctx.save();
    ctx.globalAlpha = 0.55;
    ctx.beginPath();
    ctx.arc(nx, ny, cr, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.fill();
    ctx.globalAlpha = 1;
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 1.5 * scale;
    ctx.stroke();
    ctx.restore();

    const lx = N > 1 ? pcx + (spread + cr * 0.45) * Math.cos(angles[i]) : pcx;
    const ly = N > 1 ? pcy + (spread + cr * 0.45) * Math.sin(angles[i]) : pcy;
    const fs = autoFontSize(nodes[i].text, cr * 1.2, cr * 0.7) * scale;
    ctx.font = `${fs}px sans-serif`;
    drawText(ctx, nodes[i].text, lx, ly, cr * 1.2, cr * 0.7, fs, '#333');
  }
}

function renderMatrix(ctx, nodes, x, y, w, h, themeColors, scale) {
  const grid = [
    nodes[0] || { text: '' }, nodes[1] || { text: '' },
    nodes[2] || { text: '' }, nodes[3] || { text: '' },
  ];
  const cellW = w / 2;
  const cellH = h / 2;
  const pad = 8 * scale;

  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 2; col++) {
      const idx = row * 2 + col;
      const cx2 = x + col * cellW + pad;
      const cy2 = y + row * cellH + pad;
      const cw2 = cellW - pad * 2;
      const ch2 = cellH - pad * 2;
      const color = nodeColor(idx, themeColors);
      fillRoundRect(ctx, cx2, cy2, cw2, ch2, 8 * scale, color, null);
      const fs = autoFontSize(grid[idx].text, cw2 - 10, ch2) * scale;
      ctx.font = `${fs}px sans-serif`;
      drawText(ctx, grid[idx].text, cx2 + cw2 / 2, cy2 + ch2 / 2, cw2 - 10, ch2, fs);
    }
  }
}

function renderList(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (!nodes.length) return;
  const N = nodes.length;
  const itemH = h / N;
  const dotR  = Math.min(itemH * 0.2, 14 * scale);
  const pad   = dotR * 3;

  for (let i = 0; i < N; i++) {
    const iy  = y + i * itemH;
    const cy2 = iy + itemH / 2;
    const color = nodeColor(i, themeColors);

    // Dot
    ctx.save();
    ctx.beginPath();
    ctx.arc(x + dotR, cy2, dotR, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.fill();
    ctx.restore();

    // Background bar
    fillRoundRect(ctx, x + pad, iy + itemH * 0.1, w - pad - 4 * scale, itemH * 0.8,
                  4 * scale, color + '22', null);

    // Text
    const fs = autoFontSize(nodes[i].text, w - pad - 20 * scale, itemH * 0.7) * scale;
    ctx.font = `${fs}px sans-serif`;
    ctx.fillStyle = '#333';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillText(nodes[i].text, x + pad + 8 * scale, cy2);
  }
}

function renderRelationship(ctx, nodes, x, y, w, h, themeColors, scale) {
  if (nodes.length < 2) { renderList(ctx, nodes, x, y, w, h, themeColors, scale); return; }
  const center = nodes[Math.floor(nodes.length / 2)];
  const left   = nodes.slice(0, Math.floor(nodes.length / 2));
  const right  = nodes.slice(Math.floor(nodes.length / 2) + 1);

  const pcx = x + w / 2;
  const pcy = y + h / 2;
  const midR = Math.min(w * 0.14, h * 0.25);

  // Center bubble
  const c0 = nodeColor(Math.floor(nodes.length / 2), themeColors);
  ctx.save();
  ctx.beginPath();
  ctx.arc(pcx, pcy, midR, 0, Math.PI * 2);
  ctx.fillStyle = c0;
  ctx.fill();
  ctx.restore();
  const cfs = autoFontSize(center.text, midR * 1.6, midR * 1.2) * scale;
  ctx.font = `bold ${cfs}px sans-serif`;
  drawText(ctx, center.text, pcx, pcy, midR * 2, midR * 1.8, cfs);

  const rowH = h / Math.max(left.length, right.length, 1);
  const boxW = w * 0.3;
  const boxH = rowH * 0.7;

  const drawSide = (group, side) => {
    const bx = side === 'left' ? x : x + w - boxW;
    const arrowX2 = side === 'left' ? pcx - midR : pcx + midR;
    for (let i = 0; i < group.length; i++) {
      const by = y + rowH * i + (rowH - boxH) / 2;
      const color = nodeColor(i + (side === 'right' ? left.length + 1 : 0), themeColors);
      fillRoundRect(ctx, bx, by, boxW, boxH, 5 * scale, color, null);
      const fs = autoFontSize(group[i].text, boxW - 8, boxH) * scale;
      ctx.font = `${fs}px sans-serif`;
      drawText(ctx, group[i].text, bx + boxW / 2, by + boxH / 2, boxW - 8, boxH, fs);
      const arrowX1 = side === 'left' ? bx + boxW : bx;
      drawArrow(ctx, arrowX1, by + boxH / 2, arrowX2, pcy, color + '88', 1.5 * scale);
    }
  };

  drawSide(left, 'left');
  drawSide(right, 'right');
}

// ── Fallback: generic node grid ──────────────────────────────────────────────

function renderFallback(ctx, nodes, x, y, w, h, themeColors, scale) {
  const N = nodes.length;
  if (!N) return;
  const cols = Math.ceil(Math.sqrt(N));
  const rows = Math.ceil(N / cols);
  const cellW = w / cols;
  const cellH = h / rows;
  const pad   = Math.min(cellW, cellH) * 0.1;

  for (let i = 0; i < N; i++) {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const bx = x + col * cellW + pad;
    const by = y + row * cellH + pad;
    const bw = cellW - pad * 2;
    const bh = cellH - pad * 2;
    const color = nodeColor(i, themeColors);
    fillRoundRect(ctx, bx, by, bw, bh, 8 * scale, color, null);
    const light = lighten(color);
    // Subtle gradient shine
    ctx.save();
    const grad = ctx.createLinearGradient(bx, by, bx, by + bh * 0.5);
    grad.addColorStop(0, 'rgba(255,255,255,0.18)');
    grad.addColorStop(1, 'rgba(255,255,255,0)');
    ctx.fillStyle = grad;
    ctx.beginPath();
    ctx.roundRect?.(bx, by, bw, bh, 8 * scale);
    ctx.fill();
    ctx.restore();

    const fs = autoFontSize(nodes[i].text, bw - 12 * scale, bh) * scale;
    ctx.font = `${fs}px sans-serif`;
    drawText(ctx, nodes[i].text, bx + bw / 2, by + bh / 2, bw - 12 * scale, bh, fs);
  }
}

// ── Main entry ───────────────────────────────────────────────────────────────

/**
 * Render a SmartArt diagram.
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {Document} dataDoc    — ppt/diagrams/data*.xml parsed
 * @param {Document} layoutDoc  — ppt/diagrams/layout*.xml parsed (may be null)
 * @param {number}   x, y, w, h — bounding box in canvas pixels
 * @param {object}   themeColors
 * @param {number}   scale
 */
export function renderSmartArt(ctx, dataDoc, layoutDoc, x, y, w, h, themeColors, scale) {
  const nodes = readNodes(dataDoc);
  const layout = detectLayout(layoutDoc);
  const pad = 16 * scale;

  ctx.save();
  // Subtle background
  ctx.fillStyle = '#f7f9fc';
  ctx.strokeStyle = '#e0e4ec';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.roundRect?.(x, y, w, h, 6 * scale) || ctx.rect(x, y, w, h);
  ctx.fill();
  ctx.stroke();

  const ix = x + pad;
  const iy = y + pad;
  const iw = w - pad * 2;
  const ih = h - pad * 2;

  if (!nodes.length) {
    ctx.fillStyle = '#aaa';
    ctx.font = `${12 * scale}px sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('SmartArt Diagram', x + w / 2, y + h / 2);
    ctx.restore();
    return;
  }

  switch (layout) {
    case 'process':      renderProcess(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'cycle':        renderCycle(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'hierarchy':    renderHierarchy(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'radial':       renderRadial(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'pyramid':      renderPyramid(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'funnel':       renderFunnel(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'venn':         renderVenn(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'matrix':       renderMatrix(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'list':         renderList(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    case 'relationship': renderRelationship(ctx, nodes, ix, iy, iw, ih, themeColors, scale); break;
    default:             renderFallback(ctx, nodes, ix, iy, iw, ih, themeColors, scale);
  }

  ctx.restore();
}
