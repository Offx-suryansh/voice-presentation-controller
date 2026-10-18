/**
 * colors.js — OOXML colour resolution: all 6 colour types, all transforms.
 */

import { g1, attrInt, clamp } from './utils.js';

// ── Hex / RGB / HLS conversion ──────────────────────────────────────────────

export function hexToRgb(hex) {
  const h = hex.replace('#', '');
  return {
    r: parseInt(h.slice(0, 2), 16),
    g: parseInt(h.slice(2, 4), 16),
    b: parseInt(h.slice(4, 6), 16),
  };
}

export function rgbToHex(r, g, b) {
  return '#' + [r, g, b]
    .map(v => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0'))
    .join('');
}

export function rgbToHls(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
      case g: h = ((b - r) / d + 2) / 6; break;
      case b: h = ((r - g) / d + 4) / 6; break;
    }
  }
  return { h, l, s };
}

export function hlsToRgb(h, l, s) {
  if (s === 0) {
    const v = Math.round(l * 255);
    return { r: v, g: v, b: v };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue2rgb = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1/6) return p + (q - p) * 6 * t;
    if (t < 1/2) return q;
    if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
    return p;
  };
  return {
    r: Math.round(hue2rgb(p, q, h + 1/3) * 255),
    g: Math.round(hue2rgb(p, q, h)       * 255),
    b: Math.round(hue2rgb(p, q, h - 1/3) * 255),
  };
}

// ── Preset colours (CSS colour names → hex) ─────────────────────────────────

const PRESET_COLORS = {
  black:'000000', white:'FFFFFF', red:'FF0000', green:'008000', blue:'0000FF',
  yellow:'FFFF00', cyan:'00FFFF', magenta:'FF00FF', orange:'FFA500',
  purple:'800080', pink:'FFC0CB', brown:'A52A2A', gray:'808080', grey:'808080',
  navy:'000080', teal:'008080', maroon:'800000', olive:'808000', lime:'00FF00',
  aqua:'00FFFF', fuchsia:'FF00FF', silver:'C0C0C0', coral:'FF7F50',
  salmon:'FA8072', gold:'FFD700', khaki:'F0E68C', lavender:'E6E6FA',
  beige:'F5F5DC', ivory:'FFFFF0', mintcream:'F5FFFA', azure:'F0FFFF',
  aliceblue:'F0F8FF', ghostwhite:'F8F8FF', darkred:'8B0000',
  darkgreen:'006400', darkblue:'00008B', darkcyan:'008B8B',
  darkmagenta:'8B008B', darkorange:'FF8C00', darkgray:'A9A9A9',
  darkgrey:'A9A9A9', lightgray:'D3D3D3', lightgrey:'D3D3D3',
  lightblue:'ADD8E6', lightgreen:'90EE90', lightpink:'FFB6C1',
  lightyellow:'FFFFE0', lightcyan:'E0FFFF', deepskyblue:'00BFFF',
  royalblue:'4169E1', steelblue:'4682B4', skyblue:'87CEEB',
  dodgerblue:'1E90FF', cornflowerblue:'6495ED', mediumblue:'0000CD',
  midnightblue:'191970', indigo:'4B0082', slateblue:'6A5ACD',
  blueviolet:'8A2BE2', mediumpurple:'9370DB', orchid:'DA70D6',
  violet:'EE82EE', plum:'DDA0DD', thistle:'D8BFD8', hotpink:'FF69B4',
  deeppink:'FF1493', crimson:'DC143C', firebrick:'B22222', tomato:'FF6347',
  orangered:'FF4500', darkorange2:'FF8C00', chocolate:'D2691E',
  saddlebrown:'8B4513', sienna:'A0522D', tan:'D2B48C', burlywood:'DEB887',
  wheat:'F5DEB3', moccasin:'FFE4B5', peachpuff:'FFDAB9', papayawhip:'FFEFD5',
  mistyrose:'FFE4E1', linen:'FAF0E6', oldlace:'FDF5E6', floralwhite:'FFFAF0',
  antiquewhite:'FAEBD7', bisque:'FFE4C4', blanchedalmond:'FFEBCD',
  cornsilk:'FFF8DC', lemonchiffon:'FFFACD', honeydew:'F0FFF0',
  palegreen:'98FB98', lightseagreen:'20B2AA', mediumseagreen:'3CB371',
  seagreen:'2E8B57', forestgreen:'228B22', yellowgreen:'9ACD32',
  olivedrab:'6B8E23', greenyellow:'ADFF2F', chartreuse:'7FFF00',
  springgreen:'00FF7F', mediumspringgreen:'00FA9A', aquamarine:'7FFFD4',
  turquoise:'40E0D0', mediumturquoise:'48D1CC', paleturquoise:'AFEEEE',
  cadetblue:'5F9EA0', powderblue:'B0E0E6', lightsteelblue:'B0C4DE',
  slategray:'708090', slategrey:'708090', dimgray:'696969', dimgrey:'696969',
  snow:'FFFAFA', seashell:'FFF5EE', whitesmoke:'F5F5F5', gainsboro:'DCDCDC',
};

// ── Colour transforms ────────────────────────────────────────────────────────

/**
 * Apply all OOXML colour transforms in document order.
 * `c` = { r, g, b, a }; returns the same shape.
 */
export function applyColorTransforms(c, transformEl) {
  if (!transformEl) return c;
  let { r, g, b, a = 1 } = c;

  for (const child of transformEl.children) {
    const ln = child.localName;
    const val = parseInt(child.getAttribute('val') ?? '0', 10);

    switch (ln) {
      case 'lumMod': {
        // val in 1/1000th of a percent (100000 = 100%)
        const f = val / 100000;
        const hls = rgbToHls(r, g, b);
        const rgb = hlsToRgb(hls.h, clamp(hls.l * f, 0, 1), hls.s);
        r = rgb.r; g = rgb.g; b = rgb.b;
        break;
      }
      case 'lumOff': {
        const f = val / 100000;
        const hls = rgbToHls(r, g, b);
        const rgb = hlsToRgb(hls.h, clamp(hls.l + f, 0, 1), hls.s);
        r = rgb.r; g = rgb.g; b = rgb.b;
        break;
      }
      case 'tint': {
        // Blend toward white
        const f = val / 100000;
        r = Math.round(r + (255 - r) * (1 - f));
        g = Math.round(g + (255 - g) * (1 - f));
        b = Math.round(b + (255 - b) * (1 - f));
        break;
      }
      case 'shade': {
        // Blend toward black
        const f = val / 100000;
        r = Math.round(r * f);
        g = Math.round(g * f);
        b = Math.round(b * f);
        break;
      }
      case 'satMod': {
        const f = val / 100000;
        const hls = rgbToHls(r, g, b);
        const rgb = hlsToRgb(hls.h, hls.l, clamp(hls.s * f, 0, 1));
        r = rgb.r; g = rgb.g; b = rgb.b;
        break;
      }
      case 'satOff': {
        const f = val / 100000;
        const hls = rgbToHls(r, g, b);
        const rgb = hlsToRgb(hls.h, hls.l, clamp(hls.s + f, 0, 1));
        r = rgb.r; g = rgb.g; b = rgb.b;
        break;
      }
      case 'hueMod': {
        const f = val / 100000;
        const hls = rgbToHls(r, g, b);
        const rgb = hlsToRgb((hls.h * f) % 1, hls.l, hls.s);
        r = rgb.r; g = rgb.g; b = rgb.b;
        break;
      }
      case 'hueOff': {
        const f = val / 21600000; // 60000ths of a degree, full circle = 21600000
        const hls = rgbToHls(r, g, b);
        const rgb = hlsToRgb(((hls.h + f) % 1 + 1) % 1, hls.l, hls.s);
        r = rgb.r; g = rgb.g; b = rgb.b;
        break;
      }
      case 'alpha': {
        a = val / 100000;
        break;
      }
      case 'alphaOff': {
        a = clamp(a + val / 100000, 0, 1);
        break;
      }
      case 'alphaMod': {
        a = clamp(a * val / 100000, 0, 1);
        break;
      }
      case 'inv': {
        r = 255 - r; g = 255 - g; b = 255 - b;
        break;
      }
      case 'gray': {
        const lum = Math.round(0.2126 * r + 0.7152 * g + 0.0722 * b);
        r = g = b = lum;
        break;
      }
      case 'comp': {
        const hls = rgbToHls(r, g, b);
        const rgb = hlsToRgb((hls.h + 0.5) % 1, hls.l, hls.s);
        r = rgb.r; g = rgb.g; b = rgb.b;
        break;
      }
    }
  }

  return { r: clamp(r, 0, 255), g: clamp(g, 0, 255), b: clamp(b, 0, 255), a };
}

/** Resolve an OOXML colour element to { r, g, b, a } */
export function resolveColorElement(colorEl, themeColors) {
  if (!colorEl) return null;
  const ln = colorEl.localName;

  let rgb = null;
  let a = 1;

  if (ln === 'srgbClr') {
    const val = colorEl.getAttribute('val') || '';
    if (val.length >= 6) rgb = hexToRgb(val);
  } else if (ln === 'schemeClr') {
    const schemeVal = colorEl.getAttribute('val') || '';
    // Map scheme names to base colours via themeColors
    const key = schemeVal; // e.g. "dk1", "accent1", "tx1", "bg1" …
    const hex = themeColors?.[key];
    if (hex) rgb = hexToRgb(hex);
    else rgb = { r: 0, g: 0, b: 0 }; // fallback black
  } else if (ln === 'prstClr') {
    const prstVal = (colorEl.getAttribute('val') || '').toLowerCase();
    const hex = PRESET_COLORS[prstVal];
    if (hex) rgb = hexToRgb(hex);
  } else if (ln === 'sysClr') {
    // System colour — use lastClr if available, else fallback
    const lastClr = colorEl.getAttribute('lastClr');
    if (lastClr?.length >= 6) rgb = hexToRgb(lastClr);
    else rgb = { r: 0, g: 0, b: 0 };
  } else if (ln === 'hslClr') {
    const h = attrInt(colorEl, 'hue', 0) / 21600000;
    const s = attrInt(colorEl, 'sat', 0) / 100000;
    const l = attrInt(colorEl, 'lum', 0) / 100000;
    rgb = hlsToRgb(h, l, s);
  } else if (ln === 'scRgbClr') {
    // Linear-light (0–100000)
    const r2 = attrInt(colorEl, 'r', 0) / 100000;
    const g2 = attrInt(colorEl, 'g', 0) / 100000;
    const b2 = attrInt(colorEl, 'b', 0) / 100000;
    rgb = {
      r: Math.round(Math.pow(r2, 1/2.2) * 255),
      g: Math.round(Math.pow(g2, 1/2.2) * 255),
      b: Math.round(Math.pow(b2, 1/2.2) * 255),
    };
  }

  if (!rgb) return null;

  // Apply any transforms (children of the colour element)
  return applyColorTransforms({ ...rgb, a }, colorEl);
}

/**
 * Return a CSS colour string from { r, g, b, a }.
 * @param {{r:number,g:number,b:number,a?:number}} c
 * @param {number} [alphaOverride]  Optional alpha (0–1) that overrides c.a
 */
export function colorToCss(c, alphaOverride) {
  if (!c) return 'transparent';
  const a = alphaOverride !== undefined ? alphaOverride : (c.a ?? 1);
  return a < 1
    ? `rgba(${c.r},${c.g},${c.b},${a.toFixed(3)})`
    : `rgb(${c.r},${c.g},${c.b})`;
}

/** Find the first recognised colour child element. */
export function findFirstColorChild(el) {
  if (!el) return null;
  const tags = ['srgbClr','schemeClr','prstClr','sysClr','hslClr','scRgbClr'];
  for (const tag of tags) {
    const child = g1(el, tag);
    if (child) return child;
  }
  return null;
}

/** Convenience: resolve a run's text colour, returning a CSS string or null. */
export function getRunColor(rPr, themeColors) {
  if (!rPr) return null;
  const solidFill = g1(rPr, 'solidFill');
  if (!solidFill) return null;
  const colorChild = findFirstColorChild(solidFill);
  const c = resolveColorElement(colorChild, themeColors);
  return c ? colorToCss(c) : null;
}

/** Resolve run color with paraDefRPr fallback. */
export function getRunColorInherited(rPr, paraDefRPr, themeColors) {
  const c1 = getRunColor(rPr, themeColors);
  if (c1) return c1;
  return getRunColor(paraDefRPr, themeColors);
}
