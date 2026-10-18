/**
 * shapes.js — Draw OOXML preset and custom geometries onto a Canvas 2D context.
 * Supports 50+ preset shapes, custom geometry paths, and parametric helpers.
 */
export function drawPresetGeom(ctx, prst, x, y, w, h, adjValues) {
  // Helper: angle in degrees to radians
  const deg = d => d * Math.PI / 180;

  // Adjustment value (0-100000 range, percentage of shape dimension)
  const adj = (idx, def = 50000) => {
    const v = adjValues && adjValues[idx] !== undefined ? adjValues[idx] : def;
    return v / 100000;
  };

  const cx = x + w / 2, cy = y + h / 2;
  const ss = Math.min(w, h); // short side

  ctx.beginPath();

  switch (prst) {
    case 'rect':
    case 'snip1Rect': // simplify to rect
      ctx.rect(x, y, w, h);
      break;

    case 'roundRect': {
      const r = adj(0, 16667) * ss * 0.5;
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + w - r, y);
      ctx.arcTo(x + w, y, x + w, y + r, r);
      ctx.lineTo(x + w, y + h - r);
      ctx.arcTo(x + w, y + h, x + w - r, y + h, r);
      ctx.lineTo(x + r, y + h);
      ctx.arcTo(x, y + h, x, y + h - r, r);
      ctx.lineTo(x, y + r);
      ctx.arcTo(x, y, x + r, y, r);
      ctx.closePath();
      break;
    }

    case 'ellipse':
    case 'oval':
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
      break;

    case 'triangle':
    case 'isoscelesTriangle':
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;

    case 'rightTriangle':
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;

    case 'parallelogram': {
      const a1 = adj(0, 25000);
      const off = w * a1;
      ctx.moveTo(x + off, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w - off, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'trapezoid': {
      const a1 = adj(0, 25000);
      const off = w * a1;
      ctx.moveTo(x + off, y);
      ctx.lineTo(x + w - off, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'diamond':
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;

    case 'pentagon': drawRegularPolygon(ctx, cx, cy, w/2, h/2, 5, -Math.PI/2); break;
    case 'hexagon': drawRegularPolygon(ctx, cx, cy, w/2, h/2, 6, 0); break;
    case 'heptagon': drawRegularPolygon(ctx, cx, cy, w/2, h/2, 7, -Math.PI/2); break;
    case 'octagon': drawRegularPolygon(ctx, cx, cy, w/2, h/2, 8, Math.PI/8); break;
    case 'decagon': drawRegularPolygon(ctx, cx, cy, w/2, h/2, 10, -Math.PI/2); break;
    case 'dodecagon': drawRegularPolygon(ctx, cx, cy, w/2, h/2, 12, 0); break;

    case 'plus':
    case 'cross': {
      const a1 = adj(0, 25000);
      const t = ss * a1;
      const bx = cx - t/2, tx = cx + t/2;
      const by = cy - t/2, ty = cy + t/2;
      ctx.moveTo(bx, y); ctx.lineTo(tx, y);
      ctx.lineTo(tx, by); ctx.lineTo(x+w, by);
      ctx.lineTo(x+w, ty); ctx.lineTo(tx, ty);
      ctx.lineTo(tx, y+h); ctx.lineTo(bx, y+h);
      ctx.lineTo(bx, ty); ctx.lineTo(x, ty);
      ctx.lineTo(x, by); ctx.lineTo(bx, by);
      ctx.closePath();
      break;
    }

    case 'star4':
      drawStar(ctx, cx, cy, w/2, h/2, 4, adj(0, 37500)); break;
    case 'star5':
      drawStar(ctx, cx, cy, w/2, h/2, 5, adj(0, 19098), -Math.PI/2); break;
    case 'star6':
      drawStar(ctx, cx, cy, w/2, h/2, 6, adj(0, 28868), 0); break;
    case 'star7':
      drawStar(ctx, cx, cy, w/2, h/2, 7, adj(0, 34601), -Math.PI/2); break;
    case 'star8':
      drawStar(ctx, cx, cy, w/2, h/2, 8, adj(0, 29289), Math.PI/8); break;
    case 'star10':
      drawStar(ctx, cx, cy, w/2, h/2, 10, adj(0, 30902), -Math.PI/2); break;
    case 'star12':
      drawStar(ctx, cx, cy, w/2, h/2, 12, adj(0, 37720), 0); break;
    case 'star16':
      drawStar(ctx, cx, cy, w/2, h/2, 16, adj(0, 37500), 0); break;
    case 'star24':
      drawStar(ctx, cx, cy, w/2, h/2, 24, adj(0, 37500), 0); break;
    case 'star32':
      drawStar(ctx, cx, cy, w/2, h/2, 32, adj(0, 37500), 0); break;

    case 'rightArrow': {
      const ah = adj(0, 50000); // arrow head width ratio
      const aw = adj(1, 50000); // arrow head height ratio
      const bh = (h - h * aw) / 2;
      const nb = (h * aw - h) / 2;
      const ax = x + w * (1 - ah);
      ctx.moveTo(x, y + bh);
      ctx.lineTo(ax, y + bh);
      ctx.lineTo(ax, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(ax, y + h);
      ctx.lineTo(ax, y + h - bh);
      ctx.lineTo(x, y + h - bh);
      ctx.closePath();
      break;
    }

    case 'leftArrow': {
      const ah = adj(0, 50000);
      const aw = adj(1, 50000);
      const bh = (h - h * aw) / 2;
      const ax = x + w * ah;
      ctx.moveTo(x + w, y + bh);
      ctx.lineTo(ax, y + bh);
      ctx.lineTo(ax, y);
      ctx.lineTo(x, cy);
      ctx.lineTo(ax, y + h);
      ctx.lineTo(ax, y + h - bh);
      ctx.lineTo(x + w, y + h - bh);
      ctx.closePath();
      break;
    }

    case 'upArrow': {
      const ah = adj(0, 50000);
      const aw = adj(1, 50000);
      const bw = (w - w * aw) / 2;
      const ay = y + h * ah;
      ctx.moveTo(x + bw, y + h);
      ctx.lineTo(x + bw, ay);
      ctx.lineTo(x, ay);
      ctx.lineTo(cx, y);
      ctx.lineTo(x + w, ay);
      ctx.lineTo(x + w - bw, ay);
      ctx.lineTo(x + w - bw, y + h);
      ctx.closePath();
      break;
    }

    case 'downArrow': {
      const ah = adj(0, 50000);
      const aw = adj(1, 50000);
      const bw = (w - w * aw) / 2;
      const ay = y + h * (1 - ah);
      ctx.moveTo(x + bw, y);
      ctx.lineTo(x + bw, ay);
      ctx.lineTo(x, ay);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x + w, ay);
      ctx.lineTo(x + w - bw, ay);
      ctx.lineTo(x + w - bw, y);
      ctx.closePath();
      break;
    }

    case 'leftRightArrow': {
      const ah = adj(0, 25000);
      const aw = adj(1, 50000);
      const bh = (h - h * aw) / 2;
      const lax = x + w * ah;
      const rax = x + w * (1 - ah);
      ctx.moveTo(x, cy);
      ctx.lineTo(lax, y);
      ctx.lineTo(lax, y + bh);
      ctx.lineTo(rax, y + bh);
      ctx.lineTo(rax, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(rax, y + h);
      ctx.lineTo(rax, y + h - bh);
      ctx.lineTo(lax, y + h - bh);
      ctx.lineTo(lax, y + h);
      ctx.closePath();
      break;
    }

    case 'upDownArrow': {
      const ah = adj(0, 25000);
      const aw = adj(1, 50000);
      const bw = (w - w * aw) / 2;
      const tay = y + h * ah;
      const bay = y + h * (1 - ah);
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, tay);
      ctx.lineTo(x + w - bw, tay);
      ctx.lineTo(x + w - bw, bay);
      ctx.lineTo(x + w, bay);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, bay);
      ctx.lineTo(x + bw, bay);
      ctx.lineTo(x + bw, tay);
      ctx.lineTo(x, tay);
      ctx.closePath();
      break;
    }

    case 'chevron': {
      const a = adj(0, 50000);
      const off = w * a;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w - off, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - off, y + h);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x + off, cy);
      ctx.closePath();
      break;
    }

    case 'pentagon5': // not standard but some files use it
    case 'homePlate': {
      const a = adj(0, 50000);
      ctx.moveTo(x, y);
      ctx.lineTo(x + w * (1 - a), y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w * (1 - a), y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'arc': {
      const stAng = adj(0, 16200000 / 60000) * Math.PI / 180; // stored as 60000ths of degree
      const swAng = adj(1, 28800000 / 60000) * Math.PI / 180;
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, Math.min(w, h) / 2, stAng, stAng + swAng);
      ctx.closePath();
      break;
    }

    case 'blockArc': {
      // outer arc minus inner arc
      const stAng = (adj(0, 0) * 360 - 90) * Math.PI / 180;
      const swAng = adj(1, 25000) * 360 * Math.PI / 180;
      const tck = adj(2, 25000) * Math.min(w, h) / 2;
      const outerR = Math.min(w, h) / 2;
      const innerR = outerR - tck;
      ctx.moveTo(cx + outerR * Math.cos(stAng), cy + outerR * Math.sin(stAng));
      ctx.arc(cx, cy, outerR, stAng, stAng + swAng);
      ctx.arc(cx, cy, Math.max(1, innerR), stAng + swAng, stAng, true);
      ctx.closePath();
      break;
    }

    case 'line':
    case 'straightConnector1':
      ctx.moveTo(x, cy);
      ctx.lineTo(x + w, cy);
      break;

    case 'bentConnector3':
    case 'bentConnector4':
    case 'bentConnector5':
    case 'elbowConnector':
      ctx.moveTo(x, y);
      ctx.lineTo(x, cy);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w, y + h);
      break;

    case 'curvedConnector3':
    case 'curvedConnector4':
      ctx.moveTo(x, y);
      ctx.bezierCurveTo(x, y + h/2, x + w, y + h/2, x + w, y + h);
      break;

    case 'heart': {
      drawHeart(ctx, x, y, w, h);
      break;
    }

    case 'lightningBolt': {
      ctx.moveTo(x + w * 0.6, y);
      ctx.lineTo(x + w * 0.2, y + h * 0.45);
      ctx.lineTo(x + w * 0.5, y + h * 0.45);
      ctx.lineTo(x + w * 0.4, y + h);
      ctx.lineTo(x + w * 0.8, y + h * 0.55);
      ctx.lineTo(x + w * 0.5, y + h * 0.55);
      ctx.closePath();
      break;
    }

    case 'moon': {
      ctx.arc(cx + w * 0.1, cy, h * 0.5, deg(-120), deg(120));
      ctx.arc(cx + w * 0.4, cy, h * 0.45, deg(120), deg(-120), true);
      ctx.closePath();
      break;
    }

    case 'noSmoking': {
      const r = Math.min(w, h) / 2;
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
      ctx.moveTo(cx - r * 0.7, cy + r * 0.7);
      ctx.lineTo(cx + r * 0.7, cy - r * 0.7);
      break;
    }

    case 'flowChartProcess':
      ctx.rect(x, y, w, h);
      break;

    case 'flowChartDecision':
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;

    case 'flowChartTerminator': {
      const r2 = h / 2;
      ctx.moveTo(x + r2, y);
      ctx.lineTo(x + w - r2, y);
      ctx.arc(x + w - r2, cy, r2, -Math.PI/2, Math.PI/2);
      ctx.lineTo(x + r2, y + h);
      ctx.arc(x + r2, cy, r2, Math.PI/2, -Math.PI/2);
      ctx.closePath();
      break;
    }

    case 'flowChartDocument': {
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h * 0.8);
      ctx.bezierCurveTo(x + w * 0.75, y + h * 0.8, x + w * 0.75, y + h, x + w * 0.5, y + h);
      ctx.bezierCurveTo(x + w * 0.25, y + h, x + w * 0.25, y + h * 0.8, x, y + h * 0.8);
      ctx.closePath();
      break;
    }

    case 'flowChartDatabase':
    case 'cylinder': {
      const ry = h * 0.1;
      ctx.moveTo(x, y + ry);
      ctx.ellipse(cx, y + ry, w/2, ry, 0, Math.PI, 0);
      ctx.lineTo(x + w, y + h - ry);
      ctx.ellipse(cx, y + h - ry, w/2, ry, 0, 0, Math.PI);
      ctx.closePath();
      break;
    }

    case 'cube': {
      const d = w * 0.15;
      ctx.moveTo(x + d, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - d);
      ctx.lineTo(x + w - d, y + h);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x, y + d);
      ctx.closePath();
      ctx.moveTo(x + d, y);
      ctx.lineTo(x + d, y + d);
      ctx.lineTo(x, y + d);
      ctx.moveTo(x + d, y + d);
      ctx.lineTo(x + w, y + d);
      break;
    }

    case 'callout1':
    case 'borderCallout1':
    case 'wedgeRectCallout': {
      ctx.rect(x, y, w, h * 0.8);
      ctx.moveTo(cx - w * 0.1, y + h * 0.8);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(cx + w * 0.1, y + h * 0.8);
      ctx.closePath();
      break;
    }

    case 'wedgeRoundRectCallout': {
      const r3 = h * 0.05;
      // Round rect
      ctx.moveTo(x + r3, y);
      ctx.lineTo(x + w - r3, y);
      ctx.arcTo(x + w, y, x + w, y + r3, r3);
      ctx.lineTo(x + w, y + h * 0.75 - r3);
      ctx.arcTo(x + w, y + h * 0.75, x + w - r3, y + h * 0.75, r3);
      ctx.lineTo(cx + w * 0.1, y + h * 0.75);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(cx - w * 0.1, y + h * 0.75);
      ctx.lineTo(x + r3, y + h * 0.75);
      ctx.arcTo(x, y + h * 0.75, x, y + h * 0.75 - r3, r3);
      ctx.lineTo(x, y + r3);
      ctx.arcTo(x, y, x + r3, y, r3);
      ctx.closePath();
      break;
    }

    case 'wedgeEllipseCallout': {
      ctx.ellipse(cx, cy - h * 0.05, w/2, h * 0.45, 0, 0, Math.PI * 2);
      ctx.moveTo(cx - w * 0.1, cy + h * 0.4);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(cx + w * 0.1, cy + h * 0.4);
      ctx.closePath();
      break;
    }

    case 'cloudCallout':
    case 'cloud': {
      // Cloud: overlapping circles for bumps, bottom arc for base
      // Use an offscreen clip region via compositing would be ideal;
      // here we approximate with a path of arcs.
      ctx.beginPath();
      const cBumps = [
        { cx: cx - w * 0.22, cy: cy - h * 0.12, r: w * 0.16 },
        { cx: cx - w * 0.07, cy: cy - h * 0.22, r: w * 0.18 },
        { cx: cx + w * 0.10, cy: cy - h * 0.22, r: w * 0.17 },
        { cx: cx + w * 0.25, cy: cy - h * 0.12, r: w * 0.15 },
        { cx: cx + w * 0.35, cy: cy + h * 0.05, r: w * 0.13 },
        { cx: cx - w * 0.30, cy: cy + h * 0.05, r: w * 0.13 },
      ];
      for (const b of cBumps) ctx.arc(b.cx, b.cy, b.r, 0, Math.PI * 2);
      // Base rectangle joining the bottoms of side bumps
      ctx.rect(cx - w * 0.35, cy, w * 0.70, h * 0.25);
      break;
    }

    case 'smileyFace': {
      ctx.arc(cx, cy, Math.min(w,h)/2, 0, Math.PI * 2);
      // Draw smile as open path - eye dots drawn with filled circles in renderSp
      break;
    }

    case 'donut': {
      const r4 = Math.min(w,h)/2;
      const ir = r4 * adj(0, 25000);
      ctx.arc(cx, cy, r4, 0, Math.PI * 2);
      ctx.arc(cx, cy, ir, Math.PI * 2, 0, true);
      break;
    }

    case 'bracketPair': {
      const r5 = w * 0.2;
      ctx.moveTo(x + r5, y);
      ctx.arcTo(x, y, x, y + r5, r5);
      ctx.lineTo(x, y + h - r5);
      ctx.arcTo(x, y + h, x + r5, y + h, r5);
      ctx.moveTo(x + w - r5, y);
      ctx.arcTo(x + w, y, x + w, y + r5, r5);
      ctx.lineTo(x + w, y + h - r5);
      ctx.arcTo(x + w, y + h, x + w - r5, y + h, r5);
      break;
    }

    case 'bracePair': {
      const r6 = h * 0.15;
      // Left brace
      ctx.moveTo(cx - w*0.35, y);
      ctx.bezierCurveTo(cx - w*0.45, y, cx - w*0.45, y, cx - w*0.45, y + r6);
      ctx.lineTo(cx - w*0.45, cy - r6);
      ctx.bezierCurveTo(cx - w*0.45, cy, cx - w*0.5, cy, cx - w*0.5, cy);
      ctx.bezierCurveTo(cx - w*0.5, cy, cx - w*0.45, cy, cx - w*0.45, cy + r6);
      ctx.lineTo(cx - w*0.45, y + h - r6);
      ctx.bezierCurveTo(cx - w*0.45, y + h, cx - w*0.35, y + h, cx - w*0.35, y + h);
      break;
    }

    case 'irregularSeal1':
    case 'irregularSeal2': {
      // Star-like jagged shape
      drawStar(ctx, cx, cy, w/2, h/2, 12, adj(0, 42533), 0);
      break;
    }

    case 'accentCallout1':
    case 'accentCallout2':
    case 'calloutWedgeRect':
      ctx.rect(x, y, w, h);
      break;

    case 'flowChartAlternateProcess': {
      const r7 = h * 0.15;
      ctx.moveTo(x + r7, y);
      ctx.lineTo(x + w - r7, y);
      ctx.arcTo(x + w, y, x + w, y + r7, r7);
      ctx.lineTo(x + w, y + h - r7);
      ctx.arcTo(x + w, y + h, x + w - r7, y + h, r7);
      ctx.lineTo(x + r7, y + h);
      ctx.arcTo(x, y + h, x, y + h - r7, r7);
      ctx.lineTo(x, y + r7);
      ctx.arcTo(x, y, x + r7, y, r7);
      ctx.closePath();
      break;
    }

    case 'flowChartConnector':
      ctx.arc(cx, cy, Math.min(w,h)/2, 0, Math.PI*2);
      break;

    case 'flowChartInputOutput': {
      const off2 = w * 0.2;
      ctx.moveTo(x + off2, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w - off2, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'flowChartPredefinedProcess': {
      ctx.rect(x, y, w, h);
      ctx.moveTo(x + w*0.1, y);
      ctx.lineTo(x + w*0.1, y + h);
      ctx.moveTo(x + w*0.9, y);
      ctx.lineTo(x + w*0.9, y + h);
      break;
    }

    case 'flowChartManualInput': {
      ctx.moveTo(x, y + h * 0.2);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'flowChartPreparation': {
      const off3 = w * 0.2;
      ctx.moveTo(x + off3, y);
      ctx.lineTo(x + w - off3, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - off3, y + h);
      ctx.lineTo(x + off3, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;
    }

    case 'ribbon':
    case 'ribbon2': {
      const notchH = h * 0.3;
      ctx.moveTo(x, y + notchH/2);
      ctx.lineTo(x + w*0.1, y);
      ctx.lineTo(x + w*0.1, y + h - notchH);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x + w*0.5, y + h - notchH);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x + w - w*0.1, y + h - notchH);
      ctx.lineTo(x + w - w*0.1, y);
      ctx.lineTo(x + w, y + notchH/2);
      ctx.lineTo(x + w*0.5, y + notchH);
      ctx.closePath();
      break;
    }

    case 'ellipseRibbon':
    case 'ellipseRibbon2': {
      ctx.ellipse(cx, cy, w/2, h/2, 0, 0, Math.PI * 2);
      break;
    }

    case 'teardrop': {
      const tr = Math.min(w, h) * 0.45;
      ctx.arc(cx, cy + tr, tr, -Math.PI, 0);
      ctx.bezierCurveTo(cx + tr, cy + tr - tr*0.55, cx + tr*0.1, cy - h*0.4, cx, y);
      ctx.bezierCurveTo(cx - tr*0.1, cy - h*0.4, cx - tr, cy + tr - tr*0.55, cx - tr, cy + tr);
      ctx.closePath();
      break;
    }

    default:
      // Unknown shape — draw rectangle
      ctx.rect(x, y, w, h);
      return false;
  }
  return true;
}

export function drawRegularPolygon(ctx, cx, cy, rx, ry, n, startAngle = 0) {
  for (let i = 0; i < n; i++) {
    const angle = startAngle + (i / n) * Math.PI * 2;
    const px = cx + rx * Math.cos(angle);
    const py = cy + ry * Math.sin(angle);
    if (i === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  }
  ctx.closePath();
}

export function drawStar(ctx, cx, cy, rx, ry, points, innerRatio = 0.5, startAngle = -Math.PI/2) {
  for (let i = 0; i < points * 2; i++) {
    const angle = startAngle + (i / (points * 2)) * Math.PI * 2;
    const isInner = i % 2 === 1;
    const r_x = isInner ? rx * innerRatio : rx;
    const r_y = isInner ? ry * innerRatio : ry;
    const px = cx + r_x * Math.cos(angle);
    const py = cy + r_y * Math.sin(angle);
    if (i === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  }
  ctx.closePath();
}

export function drawHeart(ctx, x, y, w, h) {
  ctx.beginPath();
  const tx = x + w / 2;
  const topY = y + h * 0.25;
  ctx.moveTo(tx, y + h * 0.9);
  // Left side
  ctx.bezierCurveTo(
    x - w * 0.1, y + h * 0.6,
    x - w * 0.1, topY - h * 0.05,
    tx - w * 0.25, topY - h * 0.1
  );
  ctx.bezierCurveTo(
    tx - w * 0.5 + w * 0.05, y + h * 0.05,
    tx - w * 0.03, y + h * 0.05,
    tx, topY - h * 0.15
  );
  // Right side
  ctx.bezierCurveTo(
    tx + w * 0.03, y + h * 0.05,
    tx + w * 0.5 - w * 0.05, y + h * 0.05,
    tx + w * 0.25, topY - h * 0.1
  );
  ctx.bezierCurveTo(
    x + w * 1.1, topY - h * 0.05,
    x + w * 1.1, y + h * 0.6,
    tx, y + h * 0.9
  );
  ctx.closePath();
}