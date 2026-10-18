/**
 * charts.js — Full OOXML chart renderer.
 *
 * Supports: bar (horizontal), column (vertical), line, pie, doughnut,
 *           area, scatter — each in clustered / stacked / percentStacked.
 *
 * Chart data lives in ppt/charts/chart1.xml and is referenced from
 * a slide via a relationship ID on the <c:chart> element inside
 * <a:graphicData uri="…/chart">.
 */

import { g1, gtn, attr, attrInt } from './utils.js';
import { resolveColorElement, findFirstColorChild, colorToCss } from './colors.js';

// ── Default palette (Office theme accent colours + extras) ──────────────────
const DEFAULT_PALETTE = [
  '#4472C4', '#ED7D31', '#A9D18E', '#FF0000',
  '#FFC000', '#5B9BD5', '#70AD47', '#C00000',
  '#7030A0', '#00B0F0', '#FF7F00', '#9E480E',
];

// ── XML helpers specific to chart namespace ─────────────────────────────────

function cv(el) {
  // <c:v> text content → number if parseable
  const t = el?.textContent?.trim();
  if (t === undefined || t === null || t === '') return null;
  const n = parseFloat(t);
  return isNaN(n) ? t : n;
}

/** Collect ordered <c:pt> values from a cache (strCache or numCache). */
function readCache(cacheEl) {
  if (!cacheEl) return [];
  const count = attrInt(g1(cacheEl, 'ptCount'), 'val', 0);
  const result = new Array(count).fill(null);
  for (const pt of gtn(cacheEl, 'pt')) {
    const idx = attrInt(pt, 'idx', 0);
    result[idx] = cv(g1(pt, 'v'));
  }
  return result;
}

/** Read series name from <c:tx>. */
function seriesName(ser) {
  const tx = g1(ser, 'tx');
  if (!tx) return null;
  // Inline string
  const v = g1(tx, 'v');
  if (v) return v.textContent.trim();
  // strRef cache
  const strCache = g1(tx, 'strCache');
  if (strCache) {
    const pt = g1(strCache, 'pt');
    const vEl = pt ? g1(pt, 'v') : null;
    return vEl ? vEl.textContent.trim() : null;
  }
  return null;
}

/** Read category labels from <c:cat> or <c:xVal>. */
function readCategories(ser) {
  const catEl = g1(ser, 'cat') || g1(ser, 'xVal');
  if (!catEl) return [];
  return readCache(g1(catEl, 'strCache') || g1(catEl, 'numCache'));
}

/** Read numeric values from <c:val> or <c:yVal>. */
function readValues(ser) {
  const valEl = g1(ser, 'val') || g1(ser, 'yVal');
  if (!valEl) return [];
  return readCache(g1(valEl, 'numCache'));
}

/** Read per-series colour from spPr > solidFill, or fall back to index. */
function seriesColor(ser, idx, themeColors) {
  const spPr = g1(ser, 'spPr');
  if (spPr) {
    const solidFill = g1(spPr, 'solidFill');
    if (solidFill) {
      const colorChild = findFirstColorChild(solidFill);
      const c = resolveColorElement(colorChild, themeColors);
      if (c) return colorToCss(c);
    }
  }
  // Use theme accent colours in order
  const accentKey = `accent${(idx % 6) + 1}`;
  if (themeColors[accentKey]) {
    const rgb = themeColors[accentKey];
    return `#${rgb}`;
  }
  return DEFAULT_PALETTE[idx % DEFAULT_PALETTE.length];
}

/** Per data-point override colour. Returns map: index → css colour string. */
function dataPointColors(ser, themeColors) {
  const map = {};
  for (const dPt of gtn(ser, 'dPt')) {
    const idx = attrInt(g1(dPt, 'idx'), 'val', -1);
    if (idx < 0) continue;
    const spPr = g1(dPt, 'spPr');
    if (!spPr) continue;
    const solidFill = g1(spPr, 'solidFill');
    if (!solidFill) continue;
    const colorChild = findFirstColorChild(solidFill);
    const c = resolveColorElement(colorChild, themeColors);
    if (c) map[idx] = colorToCss(c);
  }
  return map;
}

// ── Chart bounds ─────────────────────────────────────────────────────────────

function chartBounds(cx, cy, cw, ch, opts = {}) {
  const {
    padL = 0.12, padR = 0.06, padT = 0.08, padB = 0.10,
    legendH = 0.10, hasLegend = true, hasTitle = false,
  } = opts;
  const tOffset = hasTitle ? ch * 0.08 : 0;
  const lOffset = hasLegend ? ch * legendH : 0;
  return {
    x:  cx + cw * padL,
    y:  cy + ch * padT + tOffset,
    w:  cw * (1 - padL - padR),
    h:  ch * (1 - padT - padB) - lOffset - tOffset,
    legendY: cy + ch * (1 - legendH * 0.7),
  };
}

// ── Drawing primitives ───────────────────────────────────────────────────────

function roundRect(ctx, x, y, w, h, r = 3) {
  if (w < 0) { x += w; w = -w; }
  if (h < 0) { y += h; h = -h; }
  r = Math.min(r, Math.abs(w) / 2, Math.abs(h) / 2);
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

function drawAxisLine(ctx, x1, y1, x2, y2, color = '#999', width = 0.7) {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  ctx.setLineDash([]);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  ctx.restore();
}

function drawGridLine(ctx, x1, y1, x2, y2) {
  ctx.save();
  ctx.strokeStyle = 'rgba(0,0,0,0.10)';
  ctx.lineWidth = 0.5;
  ctx.setLineDash([3, 3]);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  ctx.setLineDash([]);
  ctx.restore();
}

function niceStep(range, targetTicks = 5) {
  const rough = range / targetTicks;
  const mag = Math.pow(10, Math.floor(Math.log10(rough)));
  const norm = rough / mag;
  let step;
  if      (norm < 1.5) step = 1;
  else if (norm < 3.5) step = 2;
  else if (norm < 7.5) step = 5;
  else                 step = 10;
  return step * mag;
}

function calcAxisRange(values, forceZero = true) {
  const flat = values.filter(v => typeof v === 'number' && isFinite(v));
  if (!flat.length) return { min: 0, max: 100, step: 20 };
  let min = forceZero ? Math.min(0, ...flat) : Math.min(...flat);
  let max = Math.max(...flat);
  if (min === max) { min -= 1; max += 1; }
  const step = niceStep(max - min);
  max = Math.ceil(max / step) * step;
  min = Math.floor(min / step) * step;
  return { min, max, step };
}

function fmtLabel(n) {
  if (typeof n !== 'number') return String(n ?? '');
  if (Math.abs(n) >= 1e9) return (n / 1e9).toFixed(1).replace(/\.0$/, '') + 'B';
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
  if (Number.isInteger(n)) return String(n);
  return n.toFixed(1);
}

// ── Legend ───────────────────────────────────────────────────────────────────

function drawLegend(ctx, cx, cy, cw, legendY, series, scale) {
  if (!series.length) return;
  const sz = Math.max(8, Math.min(14, cw * 0.025)) * scale;
  const itemW = cw / Math.max(series.length, 1);
  ctx.save();
  ctx.font = `${sz}px sans-serif`;
  ctx.textBaseline = 'middle';
  for (let i = 0; i < series.length; i++) {
    if (!series[i].name) continue;
    const lx = cx + itemW * i + sz * 0.5;
    // Colour swatch
    ctx.fillStyle = series[i].color;
    roundRect(ctx, lx, legendY - sz / 2, sz * 1.2, sz);
    ctx.fill();
    // Label
    ctx.fillStyle = '#444';
    ctx.fillText(series[i].name, lx + sz * 1.5, legendY);
  }
  ctx.restore();
}

// ─────────────────────────────────────────────────────────────────────────────
//  BAR / COLUMN CHART
// ─────────────────────────────────────────────────────────────────────────────

function renderBarChart(ctx, chartEl, cx, cy, cw, ch, themeColors, scale) {
  const barChart  = g1(chartEl, 'barChart')  || g1(chartEl, 'bar3DChart');
  const isColumn  = !barChart || attr(barChart, 'barDir', 'col') !== 'bar';
  const grouping  = attr(barChart, 'grouping', 'clustered');
  const isStacked = grouping === 'stacked' || grouping === 'percentStacked';
  const isPct     = grouping === 'percentStacked';

  const serEls = gtn(barChart, 'ser');
  const seriesData = serEls.map((s, i) => ({
    name:   seriesName(s),
    values: readValues(s),
    color:  seriesColor(s, i, themeColors),
    dptColors: dataPointColors(s, themeColors),
  }));
  if (!seriesData.length) return;

  const cats = readCategories(serEls[0]) || [];
  const numCats = Math.max(cats.length, seriesData[0]?.values.length || 0, 1);

  const b = chartBounds(cx, cy, cw, ch, { hasLegend: seriesData.length > 1 });

  // Compute axis range
  let axisVals;
  if (isStacked) {
    // sum per category
    axisVals = Array.from({ length: numCats }, (_, ci) =>
      seriesData.reduce((s, ser) => s + (ser.values[ci] || 0), 0));
    if (isPct) axisVals = axisVals.map(() => 100);
  } else {
    axisVals = seriesData.flatMap(s => s.values);
  }
  const range = calcAxisRange(axisVals);

  const fontSize = Math.max(7, Math.min(11, b.w * 0.018)) * scale;
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = 'middle';

  // ── Y axis (column) or X axis (bar) ─────────────────────────────────────
  const valueSteps = Math.round((range.max - range.min) / range.step);
  const axisLabelPad = isColumn ? b.x * 0.35 : b.h * 0.15;

  if (isColumn) {
    // Y axis: value axis on left
    ctx.save();
    ctx.textAlign = 'right';
    ctx.fillStyle = '#666';
    for (let t = 0; t <= valueSteps; t++) {
      const val = range.min + t * range.step;
      const pct = (val - range.min) / (range.max - range.min);
      const gy  = b.y + b.h - pct * b.h;
      drawGridLine(ctx, b.x, gy, b.x + b.w, gy);
      ctx.fillText(isPct ? val + '%' : fmtLabel(val), b.x - 4 * scale, gy);
    }
    ctx.restore();

    // X axis: categories
    ctx.save();
    ctx.textAlign = 'center';
    ctx.fillStyle = '#666';
    const barGroupW = b.w / numCats;
    for (let ci = 0; ci < numCats; ci++) {
      const lx = b.x + barGroupW * ci + barGroupW / 2;
      const label = String(cats[ci] ?? ci + 1);
      ctx.fillText(label, lx, b.y + b.h + fontSize * 1.5);
    }
    ctx.restore();

    drawAxisLine(ctx, b.x, b.y, b.x, b.y + b.h);
    drawAxisLine(ctx, b.x, b.y + b.h, b.x + b.w, b.y + b.h);

    // Draw bars
    const barGroupW2 = b.w / numCats;
    const gap = barGroupW2 * 0.15;
    const groupInner = barGroupW2 - gap * 2;
    const barW = isStacked ? groupInner : groupInner / seriesData.length;

    for (let ci = 0; ci < numCats; ci++) {
      const gx = b.x + barGroupW2 * ci + gap;
      let stackBase = 0;

      for (let si = 0; si < seriesData.length; si++) {
        const ser = seriesData[si];
        let val = ser.values[ci] ?? 0;
        if (isPct) {
          const total = seriesData.reduce((s, ss) => s + (ss.values[ci] || 0), 0);
          val = total ? (val / total) * 100 : 0;
        }

        const color = ser.dptColors[ci] || ser.color;
        const barH = (Math.abs(val) / (range.max - range.min)) * b.h;
        const bx   = isStacked ? gx : gx + si * barW;
        const basePct = isStacked
          ? (stackBase - range.min) / (range.max - range.min)
          : (Math.max(0, -range.min)) / (range.max - range.min);
        const baseY = b.y + b.h - basePct * b.h;
        const barY  = val >= 0 ? baseY - barH : baseY;

        ctx.save();
        ctx.fillStyle = color;
        roundRect(ctx, bx, barY, isStacked ? groupInner : barW - 1 * scale, barH, 2 * scale);
        ctx.fill();
        // subtle top shine
        const shine = ctx.createLinearGradient(bx, barY, bx, barY + barH * 0.3);
        shine.addColorStop(0, 'rgba(255,255,255,0.18)');
        shine.addColorStop(1, 'rgba(255,255,255,0)');
        ctx.fillStyle = shine;
        ctx.fill();
        ctx.restore();

        if (isStacked) stackBase += Math.abs(val);
      }
    }

  } else {
    // HORIZONTAL BAR CHART
    ctx.save();
    ctx.textAlign = 'right';
    ctx.fillStyle = '#666';
    const barGroupH = b.h / numCats;
    for (let ci = 0; ci < numCats; ci++) {
      const ly = b.y + barGroupH * ci + barGroupH / 2;
      const label = String(cats[ci] ?? ci + 1);
      ctx.fillText(label, b.x - 4 * scale, ly);
    }
    ctx.restore();

    ctx.save();
    ctx.textAlign = 'center';
    ctx.fillStyle = '#666';
    for (let t = 0; t <= valueSteps; t++) {
      const val = range.min + t * range.step;
      const pct = (val - range.min) / (range.max - range.min);
      const gx  = b.x + pct * b.w;
      drawGridLine(ctx, gx, b.y, gx, b.y + b.h);
      ctx.fillText(isPct ? val + '%' : fmtLabel(val), gx, b.y + b.h + fontSize * 1.5);
    }
    ctx.restore();

    drawAxisLine(ctx, b.x, b.y + b.h, b.x + b.w, b.y + b.h);
    drawAxisLine(ctx, b.x, b.y, b.x, b.y + b.h);

    const gap = barGroupH * 0.15;
    const groupInner = barGroupH - gap * 2;
    const barH = isStacked ? groupInner : groupInner / seriesData.length;
    const zeroX = b.x + (-range.min / (range.max - range.min)) * b.w;

    for (let ci = 0; ci < numCats; ci++) {
      const gy = b.y + barGroupH * ci + gap;
      let stackBase = 0;

      for (let si = 0; si < seriesData.length; si++) {
        const ser = seriesData[si];
        let val = ser.values[ci] ?? 0;
        if (isPct) {
          const total = seriesData.reduce((s, ss) => s + (ss.values[ci] || 0), 0);
          val = total ? (val / total) * 100 : 0;
        }
        const color = ser.dptColors[ci] || ser.color;
        const barW  = (Math.abs(val) / (range.max - range.min)) * b.w;
        const by    = isStacked ? gy : gy + si * barH;
        const baseX = isStacked ? zeroX + (stackBase / (range.max - range.min)) * b.w : zeroX;
        const bx    = val >= 0 ? baseX : baseX - barW;

        ctx.save();
        ctx.fillStyle = color;
        roundRect(ctx, bx, by, barW, isStacked ? groupInner : barH - 1 * scale, 2 * scale);
        ctx.fill();
        ctx.restore();

        if (isStacked) stackBase += Math.abs(val);
      }
    }
  }

  if (seriesData.length > 1) {
    drawLegend(ctx, cx, cy, cw, b.legendY, seriesData, scale);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  LINE CHART
// ─────────────────────────────────────────────────────────────────────────────

function renderLineChart(ctx, chartEl, cx, cy, cw, ch, themeColors, scale) {
  const lineChart = g1(chartEl, 'lineChart') || g1(chartEl, 'line3DChart');
  const serEls = gtn(lineChart, 'ser');
  const seriesData = serEls.map((s, i) => ({
    name:   seriesName(s),
    values: readValues(s),
    color:  seriesColor(s, i, themeColors),
    marker: attr(g1(s, 'marker'), 'symbol', 'none') !== 'none',
    smooth: attr(g1(s, 'smooth'), 'val', '0') === '1',
  }));
  if (!seriesData.length) return;

  const cats = readCategories(serEls[0]) || [];
  const numCats = Math.max(cats.length, seriesData[0]?.values.length || 0, 1);
  const b = chartBounds(cx, cy, cw, ch, { hasLegend: true });
  const range = calcAxisRange(seriesData.flatMap(s => s.values));
  const fontSize = Math.max(7, Math.min(11, b.w * 0.018)) * scale;

  // Axes
  const valueSteps = Math.round((range.max - range.min) / range.step);
  ctx.save();
  ctx.textAlign = 'right';
  ctx.fillStyle = '#666';
  ctx.font = `${fontSize}px sans-serif`;
  for (let t = 0; t <= valueSteps; t++) {
    const val = range.min + t * range.step;
    const pct = (val - range.min) / (range.max - range.min);
    const gy  = b.y + b.h - pct * b.h;
    drawGridLine(ctx, b.x, gy, b.x + b.w, gy);
    ctx.fillText(fmtLabel(val), b.x - 4 * scale, gy);
  }
  ctx.restore();

  ctx.save();
  ctx.textAlign = 'center';
  ctx.fillStyle = '#666';
  ctx.font = `${fontSize}px sans-serif`;
  for (let ci = 0; ci < numCats; ci++) {
    const gx = b.x + (ci / (numCats - 1 || 1)) * b.w;
    drawGridLine(ctx, gx, b.y, gx, b.y + b.h);
    const label = String(cats[ci] ?? ci + 1);
    ctx.fillText(label, gx, b.y + b.h + fontSize * 1.5);
  }
  ctx.restore();

  drawAxisLine(ctx, b.x, b.y, b.x, b.y + b.h);
  drawAxisLine(ctx, b.x, b.y + b.h, b.x + b.w, b.y + b.h);

  // Draw each series
  for (const ser of seriesData) {
    const pts = ser.values.map((v, i) => {
      const pct = (numCats > 1) ? i / (numCats - 1) : 0.5;
      const vpct = (v - range.min) / (range.max - range.min);
      return { x: b.x + pct * b.w, y: b.y + b.h - vpct * b.h, v };
    }).filter(p => typeof p.v === 'number');

    if (pts.length < 2) continue;

    // Area fill under line
    ctx.save();
    ctx.beginPath();
    ctx.moveTo(pts[0].x, b.y + b.h);
    for (const p of pts) ctx.lineTo(p.x, p.y);
    ctx.lineTo(pts[pts.length - 1].x, b.y + b.h);
    ctx.closePath();
    ctx.fillStyle = ser.color + '22';
    ctx.fill();
    ctx.restore();

    // Line
    ctx.save();
    ctx.strokeStyle = ser.color;
    ctx.lineWidth = 2 * scale;
    ctx.lineJoin = 'round';
    ctx.lineCap = 'round';
    ctx.beginPath();
    if (ser.smooth && pts.length > 2) {
      ctx.moveTo(pts[0].x, pts[0].y);
      for (let i = 1; i < pts.length - 1; i++) {
        const cpx = (pts[i].x + pts[i + 1].x) / 2;
        const cpy = (pts[i].y + pts[i + 1].y) / 2;
        ctx.quadraticCurveTo(pts[i].x, pts[i].y, cpx, cpy);
      }
      ctx.lineTo(pts[pts.length - 1].x, pts[pts.length - 1].y);
    } else {
      ctx.moveTo(pts[0].x, pts[0].y);
      for (const p of pts.slice(1)) ctx.lineTo(p.x, p.y);
    }
    ctx.stroke();
    ctx.restore();

    // Markers
    const markerR = 3.5 * scale;
    for (const p of pts) {
      ctx.save();
      ctx.fillStyle = ser.color;
      ctx.strokeStyle = '#fff';
      ctx.lineWidth = 1.5 * scale;
      ctx.beginPath();
      ctx.arc(p.x, p.y, markerR, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
      ctx.restore();
    }
  }

  drawLegend(ctx, cx, cy, cw, b.legendY, seriesData, scale);
}

// ─────────────────────────────────────────────────────────────────────────────
//  AREA CHART
// ─────────────────────────────────────────────────────────────────────────────

function renderAreaChart(ctx, chartEl, cx, cy, cw, ch, themeColors, scale) {
  const areaChart = g1(chartEl, 'areaChart') || g1(chartEl, 'area3DChart');
  const grouping  = attr(areaChart, 'grouping', 'standard');
  const isPct     = grouping === 'percentStacked';
  const isStacked = grouping === 'stacked' || isPct;

  const serEls = gtn(areaChart, 'ser');
  // Reverse so first series is on top
  const seriesData = serEls.map((s, i) => ({
    name:   seriesName(s),
    values: readValues(s),
    color:  seriesColor(s, i, themeColors),
  })).reverse();
  if (!seriesData.length) return;

  const numCats = Math.max(...seriesData.map(s => s.values.length), 1);
  const b = chartBounds(cx, cy, cw, ch, { hasLegend: true });

  let maxVal = 0;
  if (isPct) {
    maxVal = 100;
  } else {
    for (let ci = 0; ci < numCats; ci++) {
      const sum = seriesData.reduce((s, ser) => s + (ser.values[ci] || 0), 0);
      maxVal = Math.max(maxVal, sum);
    }
  }
  const range = { min: 0, max: maxVal || 100, step: niceStep(maxVal || 100) };

  // Axes
  const valueSteps = Math.round((range.max - range.min) / range.step);
  const fontSize = Math.max(7, Math.min(11, b.w * 0.018)) * scale;
  ctx.save();
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textAlign = 'right';
  ctx.fillStyle = '#666';
  for (let t = 0; t <= valueSteps; t++) {
    const val = range.min + t * range.step;
    const pct = (val - range.min) / (range.max - range.min);
    const gy  = b.y + b.h - pct * b.h;
    drawGridLine(ctx, b.x, gy, b.x + b.w, gy);
    ctx.fillText(isPct ? val + '%' : fmtLabel(val), b.x - 4 * scale, gy);
  }
  ctx.restore();

  drawAxisLine(ctx, b.x, b.y, b.x, b.y + b.h);
  drawAxisLine(ctx, b.x, b.y + b.h, b.x + b.w, b.y + b.h);

  const stacks = new Array(numCats).fill(0);

  for (const ser of seriesData) {
    ctx.save();
    ctx.beginPath();
    const topPts = [];

    for (let ci = 0; ci < numCats; ci++) {
      let val = ser.values[ci] ?? 0;
      if (isPct) {
        const total = seriesData.reduce((s, ss) => s + (ss.values[ci] || 0), 0);
        val = total ? (val / total) * 100 : 0;
      }
      const base = isStacked ? stacks[ci] : 0;
      const top  = base + Math.abs(val);
      const bpct = (base - range.min) / (range.max - range.min);
      const tpct = (top  - range.min) / (range.max - range.min);
      const xpos = b.x + (ci / (numCats - 1 || 1)) * b.w;
      topPts.push({ x: xpos, y: b.y + b.h - tpct * b.h, baseY: b.y + b.h - bpct * b.h });
      if (isStacked) stacks[ci] += Math.abs(val);
    }

    // Bottom edge (previous stack or zero line)
    ctx.moveTo(topPts[0].x, topPts[0].baseY);
    for (const p of topPts) ctx.lineTo(p.x, p.baseY);
    // Top edge reversed
    for (let i = topPts.length - 1; i >= 0; i--) ctx.lineTo(topPts[i].x, topPts[i].y);
    ctx.closePath();
    ctx.fillStyle = ser.color + 'cc';
    ctx.fill();

    // Top outline
    ctx.beginPath();
    ctx.moveTo(topPts[0].x, topPts[0].y);
    for (const p of topPts.slice(1)) ctx.lineTo(p.x, p.y);
    ctx.strokeStyle = ser.color;
    ctx.lineWidth = 1.5 * scale;
    ctx.stroke();
    ctx.restore();
  }

  drawLegend(ctx, cx, cy, cw, b.legendY, [...seriesData].reverse(), scale);
}

// ─────────────────────────────────────────────────────────────────────────────
//  PIE / DOUGHNUT CHART
// ─────────────────────────────────────────────────────────────────────────────

function renderPieChart(ctx, chartEl, cx, cy, cw, ch, themeColors, scale) {
  const pieChart  = g1(chartEl, 'pieChart')  || g1(chartEl, 'pie3DChart');
  const doughnut  = g1(chartEl, 'doughnutChart');
  const chartNode = pieChart || doughnut;
  const isDoughnut = !!doughnut;

  const serEls = gtn(chartNode, 'ser');
  if (!serEls.length) return;

  const ser = serEls[0]; // pie/doughnut always 1 series
  const values = readValues(ser).map(v => (typeof v === 'number' && v > 0 ? v : 0));
  const cats   = readCategories(ser);
  const dptClrs = dataPointColors(ser, themeColors);

  const total = values.reduce((a, b) => a + b, 0);
  if (!total) return;

  // Center and radius
  const cr  = Math.min(cw, ch) * 0.38;
  const pcx = cx + cw * 0.44;
  const pcy = cy + ch * 0.50;
  const holeR = isDoughnut ? cr * 0.55 : 0;

  let startAngle = -Math.PI / 2;

  for (let i = 0; i < values.length; i++) {
    if (!values[i]) continue;
    const sweep = (values[i] / total) * Math.PI * 2;
    const color = dptClrs[i] || seriesColor(ser, i, themeColors);
    const midA  = startAngle + sweep / 2;
    // Slight explode on hover-ish effect (first slice only)
    const explode = i === 0 ? cr * 0.04 : 0;
    const eox = explode * Math.cos(midA);
    const eoy = explode * Math.sin(midA);

    ctx.save();
    ctx.beginPath();
    ctx.moveTo(pcx + eox, pcy + eoy);
    ctx.arc(pcx + eox, pcy + eoy, cr, startAngle, startAngle + sweep);
    ctx.closePath();

    ctx.fillStyle = color;
    ctx.fill();

    // Subtle shadow on each slice
    ctx.shadowColor = 'rgba(0,0,0,0.12)';
    ctx.shadowBlur = 4 * scale;
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 1.5 * scale;
    ctx.stroke();
    ctx.restore();

    // Hole
    if (holeR > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.arc(pcx, pcy, holeR, 0, Math.PI * 2);
      ctx.fillStyle = '#fff';
      ctx.fill();
      ctx.restore();
    }

    // % label on large slices
    const pct = values[i] / total;
    if (pct > 0.05) {
      const lx = pcx + eox + (cr * 0.65) * Math.cos(midA);
      const ly = pcy + eoy + (cr * 0.65) * Math.sin(midA);
      const fontSize = Math.max(8, Math.min(13, cr * 0.15)) * scale;
      ctx.save();
      ctx.fillStyle = '#fff';
      ctx.font = `bold ${fontSize}px sans-serif`;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(Math.round(pct * 100) + '%', lx, ly);
      ctx.restore();
    }

    startAngle += sweep;
  }

  // Doughnut centre label
  if (isDoughnut) {
    const fontSize = Math.max(10, Math.min(16, cr * 0.22)) * scale;
    ctx.save();
    ctx.font = `bold ${fontSize}px sans-serif`;
    ctx.fillStyle = '#333';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(fmtLabel(total), pcx, pcy);
    ctx.restore();
  }

  // Legend on right side
  const legX  = cx + cw * 0.78;
  const fontSize = Math.max(8, Math.min(12, cw * 0.022)) * scale;
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = 'middle';
  const rowH = fontSize * 1.8;
  const startY = pcy - (values.length * rowH) / 2;

  for (let i = 0; i < values.length; i++) {
    const lx = legX;
    const ly = startY + i * rowH;
    const color = dptClrs[i] || seriesColor(ser, i, themeColors);
    ctx.fillStyle = color;
    roundRect(ctx, lx, ly - fontSize * 0.5, fontSize * 1.2, fontSize);
    ctx.fill();
    ctx.fillStyle = '#444';
    const label = String(cats[i] ?? `Item ${i + 1}`);
    ctx.fillText(label, lx + fontSize * 1.5, ly);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  SCATTER CHART
// ─────────────────────────────────────────────────────────────────────────────

function renderScatterChart(ctx, chartEl, cx, cy, cw, ch, themeColors, scale) {
  const scatterChart = g1(chartEl, 'scatterChart') || g1(chartEl, 'bubbleChart');
  const serEls = gtn(scatterChart, 'ser');
  const seriesData = serEls.map((s, i) => {
    const xVals = readCache(g1(g1(s, 'xVal'), 'numCache'));
    const yVals = readCache(g1(g1(s, 'yVal'), 'numCache'));
    const bubSz  = readCache(g1(g1(s, 'bubbleSize'), 'numCache'));
    return {
      name:   seriesName(s),
      color:  seriesColor(s, i, themeColors),
      points: xVals.map((x, j) => ({ x, y: yVals[j], r: bubSz[j] }))
               .filter(p => typeof p.x === 'number' && typeof p.y === 'number'),
    };
  });
  if (!seriesData.length) return;

  const b = chartBounds(cx, cy, cw, ch, { hasLegend: seriesData.length > 1 });

  const allX = seriesData.flatMap(s => s.points.map(p => p.x));
  const allY = seriesData.flatMap(s => s.points.map(p => p.y));
  const rangeX = calcAxisRange(allX, false);
  const rangeY = calcAxisRange(allY);
  const fontSize = Math.max(7, Math.min(11, b.w * 0.018)) * scale;

  // Grid and axes
  const stepsX = Math.round((rangeX.max - rangeX.min) / rangeX.step);
  const stepsY = Math.round((rangeY.max - rangeY.min) / rangeY.step);

  ctx.font = `${fontSize}px sans-serif`;
  ctx.save();
  ctx.textAlign = 'center';
  ctx.fillStyle = '#666';
  for (let t = 0; t <= stepsX; t++) {
    const val = rangeX.min + t * rangeX.step;
    const pct = (val - rangeX.min) / (rangeX.max - rangeX.min);
    const gx  = b.x + pct * b.w;
    drawGridLine(ctx, gx, b.y, gx, b.y + b.h);
    ctx.fillText(fmtLabel(val), gx, b.y + b.h + fontSize * 1.5);
  }
  ctx.textAlign = 'right';
  for (let t = 0; t <= stepsY; t++) {
    const val = rangeY.min + t * rangeY.step;
    const pct = (val - rangeY.min) / (rangeY.max - rangeY.min);
    const gy  = b.y + b.h - pct * b.h;
    drawGridLine(ctx, b.x, gy, b.x + b.w, gy);
    ctx.fillText(fmtLabel(val), b.x - 4 * scale, gy);
  }
  ctx.restore();

  drawAxisLine(ctx, b.x, b.y, b.x, b.y + b.h);
  drawAxisLine(ctx, b.x, b.y + b.h, b.x + b.w, b.y + b.h);

  const maxR = seriesData.flatMap(s => s.points.map(p => p.r ?? 1));
  const maxBubble = Math.max(...maxR.filter(v => typeof v === 'number'), 1);
  const maxBubbleR = Math.min(b.w, b.h) * 0.06;

  for (const ser of seriesData) {
    for (const pt of ser.points) {
      const px = b.x + ((pt.x - rangeX.min) / (rangeX.max - rangeX.min)) * b.w;
      const py = b.y + b.h - ((pt.y - rangeY.min) / (rangeY.max - rangeY.min)) * b.h;
      const r  = pt.r != null ? (pt.r / maxBubble) * maxBubbleR : 4 * scale;
      ctx.save();
      ctx.fillStyle = ser.color + 'aa';
      ctx.strokeStyle = ser.color;
      ctx.lineWidth = 1 * scale;
      ctx.beginPath();
      ctx.arc(px, py, r, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
      ctx.restore();
    }
  }

  if (seriesData.length > 1) {
    drawLegend(ctx, cx, cy, cw, b.legendY, seriesData, scale);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  RADAR CHART
// ─────────────────────────────────────────────────────────────────────────────

function renderRadarChart(ctx, chartEl, cx, cy, cw, ch, themeColors, scale) {
  const radarChart = g1(chartEl, 'radarChart');
  const serEls = gtn(radarChart, 'ser');
  const seriesData = serEls.map((s, i) => ({
    name:   seriesName(s),
    values: readValues(s),
    color:  seriesColor(s, i, themeColors),
  }));
  if (!seriesData.length) return;

  const cats = readCategories(serEls[0]) || [];
  const N = cats.length || seriesData[0]?.values.length || 0;
  if (N < 3) return;

  const pcx = cx + cw * 0.50;
  const pcy = cy + ch * 0.50;
  const r   = Math.min(cw, ch) * 0.34;
  const range = calcAxisRange(seriesData.flatMap(s => s.values));
  const rings = 4;

  // Web grid
  for (let ring = 1; ring <= rings; ring++) {
    const rr = r * ring / rings;
    ctx.save();
    ctx.beginPath();
    for (let i = 0; i < N; i++) {
      const angle = (i / N) * Math.PI * 2 - Math.PI / 2;
      const px = pcx + rr * Math.cos(angle);
      const py = pcy + rr * Math.sin(angle);
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath();
    ctx.strokeStyle = 'rgba(0,0,0,0.12)';
    ctx.lineWidth = 0.5;
    ctx.stroke();
    ctx.restore();
  }

  // Spokes
  for (let i = 0; i < N; i++) {
    const angle = (i / N) * Math.PI * 2 - Math.PI / 2;
    drawAxisLine(ctx, pcx, pcy, pcx + r * Math.cos(angle), pcy + r * Math.sin(angle));
    // Category labels
    const lx = pcx + (r + 16 * scale) * Math.cos(angle);
    const ly = pcy + (r + 16 * scale) * Math.sin(angle);
    ctx.save();
    ctx.font = `${Math.max(7, 11 * scale)}px sans-serif`;
    ctx.fillStyle = '#555';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(String(cats[i] ?? i + 1), lx, ly);
    ctx.restore();
  }

  // Data polygons
  for (const ser of seriesData) {
    const pts = ser.values.map((v, i) => {
      const angle = (i / N) * Math.PI * 2 - Math.PI / 2;
      const pct = (v - range.min) / (range.max - range.min);
      return {
        x: pcx + r * pct * Math.cos(angle),
        y: pcy + r * pct * Math.sin(angle),
      };
    });

    ctx.save();
    ctx.beginPath();
    ctx.moveTo(pts[0].x, pts[0].y);
    for (const p of pts.slice(1)) ctx.lineTo(p.x, p.y);
    ctx.closePath();
    ctx.fillStyle = ser.color + '44';
    ctx.fill();
    ctx.strokeStyle = ser.color;
    ctx.lineWidth = 1.5 * scale;
    ctx.stroke();
    ctx.restore();
  }

  drawLegend(ctx, cx, cy, cw, cy + ch - 20 * scale, seriesData, scale);
}

// ─────────────────────────────────────────────────────────────────────────────
//  CHART TITLE
// ─────────────────────────────────────────────────────────────────────────────

function drawChartTitle(ctx, chartEl, cx, cy, cw, scale) {
  const titleEl = g1(g1(chartEl, 'chart'), 'title');
  if (!titleEl) return false;
  const txEl = g1(titleEl, 'tx');
  if (!txEl) return false;
  let text = '';
  for (const t of gtn(txEl, 't')) text += t.textContent;
  if (!text.trim()) return false;

  const fontSize = Math.max(10, Math.min(16, cw * 0.030)) * scale;
  ctx.save();
  ctx.font = `bold ${fontSize}px sans-serif`;
  ctx.fillStyle = '#333';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillText(text, cx + cw / 2, cy + 6 * scale);
  ctx.restore();
  return true;
}

// ─────────────────────────────────────────────────────────────────────────────
//  MAIN ENTRY POINT
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Render a chart from its parsed XML document.
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {Document}  chartDoc  — parsed chart1.xml
 * @param {number}    cx, cy, cw, ch  — bounding box on canvas (px)
 * @param {object}    themeColors     — resolved theme colours
 * @param {number}    scale           — px-per-EMU scale factor
 */
export function renderChart(ctx, chartDoc, cx, cy, cw, ch, themeColors, scale) {
  if (!chartDoc) return;
  const chartEl = chartDoc;

  // Chart background
  ctx.save();
  ctx.fillStyle = '#fff';
  roundRect(ctx, cx, cy, cw, ch, 4 * scale);
  ctx.fill();
  ctx.strokeStyle = '#ddd';
  ctx.lineWidth = 0.5;
  ctx.stroke();
  ctx.restore();

  // Title
  const hasTitle = drawChartTitle(ctx, chartEl, cx, cy, cw, scale);

  const ty = hasTitle ? cy + Math.min(ch * 0.08, 24 * scale) : cy;
  const th = ch - (ty - cy);

  const plotArea = g1(chartEl, 'plotArea');
  if (!plotArea) return;

  const b3d = g1(plotArea, 'bar3DChart') || g1(plotArea, 'barChart');
  const l3d = g1(plotArea, 'line3DChart') || g1(plotArea, 'lineChart');
  const a3d = g1(plotArea, 'area3DChart') || g1(plotArea, 'areaChart');
  const p3d = g1(plotArea, 'pie3DChart')  || g1(plotArea, 'pieChart');
  const dnut = g1(plotArea, 'doughnutChart');
  const sct  = g1(plotArea, 'scatterChart') || g1(plotArea, 'bubbleChart');
  const rdr  = g1(plotArea, 'radarChart');

  // Pass plotArea as 'chartEl' so each renderer can find its chart node
  if (b3d)  renderBarChart(ctx, plotArea, cx, ty, cw, th, themeColors, scale);
  else if (l3d)  renderLineChart(ctx, plotArea, cx, ty, cw, th, themeColors, scale);
  else if (a3d)  renderAreaChart(ctx, plotArea, cx, ty, cw, th, themeColors, scale);
  else if (dnut) renderPieChart(ctx, plotArea, cx, ty, cw, th, themeColors, scale);
  else if (p3d)  renderPieChart(ctx, plotArea, cx, ty, cw, th, themeColors, scale);
  else if (sct)  renderScatterChart(ctx, plotArea, cx, ty, cw, th, themeColors, scale);
  else if (rdr)  renderRadarChart(ctx, plotArea, cx, ty, cw, th, themeColors, scale);
  else {
    // Unknown chart type — show type name
    ctx.save();
    ctx.fillStyle = '#aaa';
    ctx.font = `${12 * scale}px sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const typeEl = plotArea.firstElementChild;
    ctx.fillText(typeEl?.localName ?? 'Chart', cx + cw / 2, cy + ch / 2);
    ctx.restore();
  }
}
