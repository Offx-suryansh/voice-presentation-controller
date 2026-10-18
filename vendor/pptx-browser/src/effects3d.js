/**
 * effects3d.js — OOXML 3D effect renderer.
 *
 * Implements sp3d (shape-level 3D) and scene3d (scene-level) effects
 * using canvas 2D gradient tricks to simulate:
 *
 *   • Bevel:      top/bottom/left/right gradient edges
 *   • Extrusion:  depth offset shadow-box
 *   • Contour:    outline with 3D colour
 *   • Camera:     perspective warp (affine approximation via ctx.transform)
 *   • Lighting:   diffuse + specular gradient overlay
 *
 * Usage:
 *   const cleanup = apply3DEffects(ctx, spPr, themeColors, x, y, w, h, scale);
 *   // ... draw shape path ...
 *   // fill the shape
 *   draw3DOverlay(ctx, sp3d, scene3d, themeColors, x, y, w, h, scale);
 *   cleanup();
 */

import { g1, attrInt, attr } from './utils.js';
import { resolveColorElement, findFirstColorChild, colorToCss } from './colors.js';

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Parse an EMU-valued attribute to canvas pixels. */
function emu(el, name, def = 0, scale = 1) {
  const v = el ? parseInt(el.getAttribute(name) || def, 10) : def;
  return v * scale;
}

/** Lum mod: darken/lighten a CSS hex colour by a factor (0–2). */
function lumMod(hex, factor) {
  if (!hex || !hex.startsWith('#')) return hex;
  let r = parseInt(hex.slice(1, 3), 16);
  let g = parseInt(hex.slice(3, 5), 16);
  let b = parseInt(hex.slice(5, 7), 16);
  r = Math.round(Math.min(255, r * factor));
  g = Math.round(Math.min(255, g * factor));
  b = Math.round(Math.min(255, b * factor));
  return `rgb(${r},${g},${b})`;
}

function hexToCss(hex) {
  return hex ? '#' + hex : null;
}

/** Read a colour element group from sp3d sub-elements like bevelT / extrusionClr. */
function read3dColor(parentEl, themeColors) {
  if (!parentEl) return null;
  const colorChild = findFirstColorChild(parentEl);
  if (!colorChild) return null;
  const c = resolveColorElement(colorChild, themeColors);
  return c ? colorToCss(c) : null;
}

// ── Bevel ─────────────────────────────────────────────────────────────────────

/**
 * Draw a bevel effect (raised or inset edges) over a filled shape.
 *
 * Reads: sp3d > bevelT, bevelB (top and bottom bevel)
 * Simulated with thin gradient strips along each edge.
 */
function drawBevel(ctx, sp3d, x, y, w, h, scale) {
  const bevelT = g1(sp3d, 'bevelT');
  const bevelB = g1(sp3d, 'bevelB');
  const bevel  = bevelT || bevelB;
  if (!bevel) return;

  const bw = Math.max(2 * scale, emu(bevel, 'w', 76200, scale));
  const bh = Math.max(2 * scale, emu(bevel, 'h', 76200, scale));
  const prst = attr(bevel, 'prst', 'circle');
  const inset = prst.toLowerCase().includes('in');

  // Clamp bevel to half shape size
  const bwx = Math.min(bw, w * 0.3);
  const bhy = Math.min(bh, h * 0.3);

  const lightColor = inset ? 'rgba(0,0,0,0.18)' : 'rgba(255,255,255,0.35)';
  const shadowColor = inset ? 'rgba(255,255,255,0.20)' : 'rgba(0,0,0,0.20)';

  // Top edge
  const gTop = ctx.createLinearGradient(x, y, x, y + bhy);
  gTop.addColorStop(0, lightColor);
  gTop.addColorStop(1, 'rgba(255,255,255,0)');
  ctx.save();
  ctx.fillStyle = gTop;
  ctx.fillRect(x, y, w, bhy);

  // Bottom edge
  const gBot = ctx.createLinearGradient(x, y + h - bhy, x, y + h);
  gBot.addColorStop(0, 'rgba(0,0,0,0)');
  gBot.addColorStop(1, shadowColor);
  ctx.fillStyle = gBot;
  ctx.fillRect(x, y + h - bhy, w, bhy);

  // Left edge
  const gLeft = ctx.createLinearGradient(x, y, x + bwx, y);
  gLeft.addColorStop(0, lightColor);
  gLeft.addColorStop(1, 'rgba(255,255,255,0)');
  ctx.fillStyle = gLeft;
  ctx.fillRect(x, y, bwx, h);

  // Right edge
  const gRight = ctx.createLinearGradient(x + w - bwx, y, x + w, y);
  gRight.addColorStop(0, 'rgba(0,0,0,0)');
  gRight.addColorStop(1, shadowColor);
  ctx.fillStyle = gRight;
  ctx.fillRect(x + w - bwx, y, bwx, h);

  ctx.restore();
}

// ── Extrusion ─────────────────────────────────────────────────────────────────

/**
 * Draw a 3D extrusion effect — a solid depth offset below/behind the shape.
 *
 * Reads: sp3d extrusionH, extrusionClr
 */
function drawExtrusion(ctx, sp3d, themeColors, x, y, w, h, scale) {
  const extH = emu(sp3d, 'extrusionH', 0, scale);
  if (extH < 1 * scale) return;

  const clrEl = g1(sp3d, 'extrusionClr') || g1(sp3d, 'contourClr');
  const color = read3dColor(clrEl, themeColors) || 'rgba(0,0,0,0.3)';

  // Depth offset — project slightly down-right
  const depth = Math.min(extH / 914400 * 72 * scale, Math.min(w, h) * 0.15);
  const dx = depth * 0.7;
  const dy = depth * 0.7;

  ctx.save();
  ctx.fillStyle = color;
  // Right face
  ctx.beginPath();
  ctx.moveTo(x + w, y);
  ctx.lineTo(x + w + dx, y + dy);
  ctx.lineTo(x + w + dx, y + h + dy);
  ctx.lineTo(x + w, y + h);
  ctx.closePath();
  ctx.fill();
  // Bottom face
  ctx.beginPath();
  ctx.moveTo(x, y + h);
  ctx.lineTo(x + dx, y + h + dy);
  ctx.lineTo(x + w + dx, y + h + dy);
  ctx.lineTo(x + w, y + h);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

// ── Contour ───────────────────────────────────────────────────────────────────

function drawContour(ctx, sp3d, themeColors, x, y, w, h, scale) {
  const contourW = emu(sp3d, 'contourW', 0, scale);
  if (contourW < 0.5) return;
  const clrEl = g1(sp3d, 'contourClr');
  const color = read3dColor(clrEl, themeColors) || '#888888';
  const cw = Math.max(0.5, contourW / 914400 * 72 * scale);

  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = cw;
  ctx.strokeRect(x, y, w, h);
  ctx.restore();
}

// ── Camera / scene3d ──────────────────────────────────────────────────────────

/**
 * Apply a camera perspective transform to the context.
 *
 * scene3d > camera: reads rot (latitude, longitude) and fov
 * Approximated as a simple skew/scale affine transform.
 *
 * Returns a cleanup function that restores the context.
 */
export function applyCamera(ctx, scene3d, x, y, w, h) {
  if (!scene3d) return () => {};

  const camera = g1(scene3d, 'camera');
  if (!camera) return () => {};

  const prst = attr(camera, 'prst', 'orthographicFront');
  const rot  = g1(camera, 'rot');

  // Only apply non-trivial camera presets
  const isOrtho = prst.toLowerCase().includes('orthographic');
  if (isOrtho && !rot) return () => {};

  // lat/lon in 60000ths of a degree
  const lat = rot ? attrInt(rot, 'lat', 0) / 60000 : 0;
  const lon = rot ? attrInt(rot, 'lon', 0) / 60000 : 0;
  const rev = rot ? attrInt(rot, 'rev', 0) / 60000 : 0;

  // Convert to radians for gentle perspective skew
  const latR = (lat * Math.PI) / 180;
  const lonR = (lon * Math.PI) / 180;

  // Max skew ±15% of dimension to keep it legible
  const skewX = Math.sin(lonR) * 0.15;
  const skewY = Math.sin(latR) * 0.08;

  ctx.save();
  ctx.transform(
    1,     skewY,
    skewX, 1,
    x * -skewX, y * -skewY  // pivot at shape origin
  );

  return () => ctx.restore();
}

// ── Lighting overlay ──────────────────────────────────────────────────────────

/**
 * Apply a lighting/specular overlay on top of a filled shape.
 *
 * scene3d > lightRig: reads dir, rig
 * Approximated as a directional radial gradient.
 */
function drawLighting(ctx, scene3d, x, y, w, h) {
  if (!scene3d) return;

  const lightRig = g1(scene3d, 'lightRig');
  if (!lightRig) return;

  const dir = attr(lightRig, 'dir', 't');
  const rig = attr(lightRig, 'rig', 'balanced');

  // Map light direction to gradient origin
  const dirMap = {
    t:  { gx: x + w / 2, gy: y },
    b:  { gx: x + w / 2, gy: y + h },
    l:  { gx: x,         gy: y + h / 2 },
    r:  { gx: x + w,     gy: y + h / 2 },
    tl: { gx: x,         gy: y },
    tr: { gx: x + w,     gy: y },
    bl: { gx: x,         gy: y + h },
    br: { gx: x + w,     gy: y + h },
  };
  const { gx, gy } = dirMap[dir] || dirMap.t;

  const intensity = rig === 'flat' ? 0.08
    : rig === 'balanced' ? 0.14
    : rig === 'sunrise' ? 0.20
    : rig === 'harsh' ? 0.28
    : 0.12;

  const radius = Math.sqrt(w * w + h * h);
  const grad = ctx.createRadialGradient(gx, gy, 0, gx, gy, radius);
  grad.addColorStop(0, `rgba(255,255,255,${intensity})`);
  grad.addColorStop(0.5, 'rgba(255,255,255,0)');
  grad.addColorStop(1, `rgba(0,0,0,${intensity * 0.5})`);

  ctx.save();
  ctx.fillStyle = grad;
  ctx.fillRect(x, y, w, h);
  ctx.restore();
}

// ── Main API ─────────────────────────────────────────────────────────────────

/**
 * Apply 3D extrusion *before* drawing the shape (so it appears behind).
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {Element}  spPr        — shape properties element (or null)
 * @param {object}   themeColors
 * @param {number}   x, y, w, h — shape bounding box in canvas pixels
 * @param {number}   scale
 * @returns {{ applyCamera: function, overlay: function }} — call overlay() after filling the shape
 */
export function setup3D(ctx, spPr, themeColors, x, y, w, h, scale) {
  if (!spPr) return { applyCamera: () => () => {}, overlay: () => {} };

  const sp3d    = g1(spPr, 'sp3d')    || g1(spPr, 'scene3d')?.parentNode && null;
  const scene3d = g1(spPr, 'scene3d');

  // Draw extrusion first (behind the shape)
  if (sp3d) {
    drawExtrusion(ctx, sp3d, themeColors, x, y, w, h, scale);
  }

  const cleanupCamera = applyCamera(ctx, scene3d, x, y, w, h);

  return {
    /** Call this *after* filling the shape to draw bevel + lighting on top. */
    overlay() {
      if (sp3d) {
        drawBevel(ctx, sp3d, x, y, w, h, scale);
        drawContour(ctx, sp3d, themeColors, x, y, w, h, scale);
      }
      if (scene3d) {
        drawLighting(ctx, scene3d, x, y, w, h);
      }
    },
    /** Call this when done to restore canvas state. */
    cleanup: cleanupCamera,
  };
}

/**
 * Simpler function: returns true if spPr contains any 3D directives.
 */
export function has3D(spPr) {
  if (!spPr) return false;
  return !!(g1(spPr, 'sp3d') || g1(spPr, 'scene3d'));
}
