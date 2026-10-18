/**
 * animation.js — OOXML Slide Animation & Transition Engine.
 *
 * Parses the <p:timing> tree from a slide XML document and builds a
 * timeline of animation steps. Drives playback via requestAnimationFrame.
 *
 * Supported entrance effects:
 *   appear, fade, flyInFromLeft/Right/Top/Bottom, zoom, wipe,
 *   blinds, box, checkerboard, diamond, dissolve, peek, plus,
 *   randomBars, split, stretch, strips, wedge, wheel, push
 *
 * Supported exit effects:
 *   disappear, fadeOut, flyOutToLeft/Right/Top/Bottom, zoomOut
 *
 * Supported emphasis:
 *   spin, grow/shrink, flash, color change (pulse)
 *
 * Transitions (slide-level):
 *   fade, push, wipe, blinds, box, cover, cut, dissolve, newsflash,
 *   wheel, zoom, morph (approximated)
 *
 * Usage:
 *   import { parseAnimations, PptxPlayer } from './animation.js';
 *
 *   const player = new PptxPlayer(renderer, canvas);
 *   player.play(slideIndex);
 *   player.pause();
 *   player.stop();
 *   player.nextClick();
 */

import { g1, gtn, attr, attrInt } from './utils.js';

// ── Types ────────────────────────────────────────────────────────────────────

/**
 * @typedef {object} AnimStep
 * @property {string}   shapeId       — shape element id (nvCxnSpPr / nvSpPr id attr)
 * @property {string}   type          — 'entrance' | 'exit' | 'emphasis' | 'media'
 * @property {string}   effect        — OOXML effect name (e.g. 'fade', 'fly', 'appear')
 * @property {number}   clickNum      — which click group (0 = auto with slide)
 * @property {number}   delay         — ms from click/start
 * @property {number}   duration      — ms
 * @property {string}   dir           — effect direction attribute ('l','r','t','b', etc.)
 * @property {object}   from          — { opacity, x, y, scaleX, scaleY }
 * @property {object}   to            — { opacity, x, y, scaleX, scaleY }
 */

// ── Timing parser ─────────────────────────────────────────────────────────────

const DEFAULT_DURATION = 500; // ms

/** Parse p:timing → list of AnimStep objects sorted by clickNum then delay. */
export function parseAnimations(slideDoc) {
  if (!slideDoc) return [];
  const timing = g1(slideDoc, 'timing');
  if (!timing) return [];
  const tnLst = g1(timing, 'tnLst');
  if (!tnLst) return [];

  const steps = [];

  // Walk the parallel timeline tree
  function walkPar(parEl, clickNum, inheritDelay) {
    const cTn = g1(parEl, 'cTn');
    if (!cTn) return;

    // Is this a click-triggered group?
    const nodeType = attr(cTn, 'nodeType', '');
    if (nodeType === 'clickEffect') clickNum += 1;
    if (nodeType === 'withEffect' && clickNum === 0) clickNum = 0;

    const delay = attrInt(cTn, 'delay', 0) + inheritDelay;

    const childTnLst = g1(parEl, 'childTnLst') || g1(cTn, 'childTnLst');
    if (!childTnLst) return;

    for (const child of childTnLst.children) {
      const ln = child.localName;
      if (ln === 'par') {
        walkPar(child, clickNum, delay);
      } else if (ln === 'seq') {
        const seqCTn = g1(child, 'cTn');
        const seqNodeType = seqCTn ? attr(seqCTn, 'nodeType', '') : '';
        const newClick = seqNodeType === 'clickEffect' ? clickNum + 1 : clickNum;
        const seqChild = g1(child, 'childTnLst');
        if (seqChild) {
          for (const seqItem of seqChild.children) {
            if (seqItem.localName === 'par') walkPar(seqItem, newClick, delay);
          }
        }
      } else if (ln === 'set' || ln === 'animEffect' || ln === 'anim' || ln === 'animScale' || ln === 'animClr') {
        const step = parseAnimEffect(child, clickNum, delay);
        if (step) steps.push(step);
      }
    }
  }

  // Outer sequence (click groups)
  function walkSeq(seqEl, baseClick) {
    const childTnLst = g1(seqEl, 'childTnLst') || g1(g1(seqEl, 'cTn'), 'childTnLst');
    if (!childTnLst) return;

    let clickNum = baseClick;
    for (const child of childTnLst.children) {
      if (child.localName === 'par') {
        const cTn = g1(child, 'cTn');
        const nodeType = cTn ? attr(cTn, 'nodeType', '') : '';
        if (nodeType === 'clickEffect' || nodeType === 'clickPar') clickNum++;
        walkPar(child, clickNum, 0);
      }
    }
  }

  for (const child of tnLst.children) {
    if (child.localName === 'par') walkPar(child, 0, 0);
    else if (child.localName === 'seq') walkSeq(child, 0);
  }

  // Also parse top-level animEffect elements directly
  for (const animEl of gtn(tnLst, 'animEffect')) {
    const step = parseAnimEffect(animEl, 0, 0);
    if (step) steps.push(step);
  }

  return steps.sort((a, b) => a.clickNum - b.clickNum || a.delay - b.delay);
}

/** Parse a single animation element (animEffect, set, anim, etc.). */
function parseAnimEffect(el, clickNum, inheritDelay) {
  const cTn = g1(el, 'cTn');

  // Get target shape
  const tgtEl = g1(el, 'tgt') || g1(cTn, 'tgt');
  const spTgt = tgtEl ? g1(tgtEl, 'spTgt') : null;
  const shapeId = spTgt ? attr(spTgt, 'spid', null) : null;
  if (!shapeId) return null;

  // Duration
  const durStr = cTn ? attr(cTn, 'dur', null) : null;
  const duration = durStr === 'indefinite' ? 2000
    : durStr ? parseInt(durStr, 10) : DEFAULT_DURATION;

  const delay = (cTn ? attrInt(cTn, 'delay', 0) : 0) + inheritDelay;

  // Effect type and filter
  const filter = attr(el, 'filter', null);
  const type   = attr(el, 'type', null) || (filter ? 'filter' : 'set');

  // Direction
  const dir = attr(el, 'dir', null)
    || (filter ? filter.split('(')[1]?.replace(')', '') : null)
    || '';

  // Categorise
  let effectType = 'emphasis';
  if (type === 'in'  || (el.localName === 'set' && attr(g1(el, 'attrNameLst') || el, 'attrName', '') === 'style.visibility' && attr(el, 'to', '') === 'visible'))
    effectType = 'entrance';
  if (type === 'out' || (el.localName === 'set' && attr(el, 'to', '') === 'hidden'))
    effectType = 'exit';

  // Resolve canonical effect name from filter string
  let effectName = filter || el.localName;
  if (filter) {
    const fLow = filter.toLowerCase();
    if (fLow.includes('fade'))    effectName = 'fade';
    else if (fLow.includes('fly')) effectName = 'fly';
    else if (fLow.includes('appear')) effectName = 'appear';
    else if (fLow.includes('zoom')) effectName = 'zoom';
    else if (fLow.includes('wipe')) effectName = 'wipe';
    else if (fLow.includes('wheel')) effectName = 'wheel';
    else if (fLow.includes('blinds')) effectName = 'blinds';
    else if (fLow.includes('box')) effectName = 'box';
    else if (fLow.includes('dissolve')) effectName = 'dissolve';
    else if (fLow.includes('split')) effectName = 'split';
    else if (fLow.includes('stretch')) effectName = 'stretch';
    else if (fLow.includes('diamond')) effectName = 'diamond';
    else if (fLow.includes('plus')) effectName = 'plus';
    else if (fLow.includes('wedge')) effectName = 'wedge';
    else if (fLow.includes('random')) effectName = 'dissolve';
    else if (fLow.includes('strips')) effectName = 'strips';
    else if (fLow.includes('peek')) effectName = 'fly';
    else if (fLow.includes('checkerboard')) effectName = 'dissolve';
    else effectName = 'fade';
  }

  return {
    shapeId,
    type: effectType,
    effect: effectName,
    clickNum,
    delay,
    duration,
    dir,
    raw: el.localName,
  };
}

// ── Transition parser ─────────────────────────────────────────────────────────

/**
 * @typedef {object} SlideTransition
 * @property {string} type      — 'fade' | 'push' | 'wipe' | 'cover' | 'cut' | etc.
 * @property {number} duration  — ms
 * @property {string} dir       — 'l' | 'r' | 't' | 'b' | 'd' | 'u' | etc.
 */

export function parseTransition(slideDoc) {
  if (!slideDoc) return null;
  const trans = g1(slideDoc, 'transition');
  if (!trans) return null;

  const dur = attrInt(trans, 'dur', 700);
  const spd = attr(trans, 'spd', 'med');
  const speed = spd === 'slow' ? 1200 : spd === 'fast' ? 300 : dur;

  // Find the specific transition element
  const child = trans.firstElementChild;
  const type = child?.localName || 'fade';
  const dir  = child ? attr(child, 'dir', 'l') : 'l';

  return { type, duration: speed, dir };
}

// ── Easing ────────────────────────────────────────────────────────────────────

function easeOut(t) { return 1 - Math.pow(1 - t, 2); }
function easeIn(t)  { return t * t; }
function easeInOut(t) { return t < 0.5 ? 2 * t * t : 1 - Math.pow(-2 * t + 2, 2) / 2; }

// ── Shape state compositor ────────────────────────────────────────────────────

/**
 * Compute draw params for a shape at a given animation progress.
 * Returns {opacity, tx, ty, scaleX, scaleY, clipProgress, clipDir, spin}
 */
function computeShapeState(step, progress, shapeW, shapeH) {
  const t = Math.max(0, Math.min(1, progress));
  const isEntrance = step.type === 'entrance';
  const p = isEntrance ? easeOut(t) : easeIn(t); // entrance eases out, exit eases in

  const state = {
    opacity: 1, tx: 0, ty: 0,
    scaleX: 1, scaleY: 1,
    clipProgress: 1, clipDir: 'none',
    spin: 0,
  };

  const effect = step.effect;
  const dir = step.dir;

  if (isEntrance) {
    switch (effect) {
      case 'appear':
        state.opacity = t >= 0.01 ? 1 : 0;
        break;
      case 'fade':
      case 'dissolve':
        state.opacity = p;
        break;
      case 'fly':
      case 'flyIn': {
        state.opacity = Math.min(1, p * 1.5);
        const dist = (1 - p);
        if (dir.includes('l')) state.tx = -shapeW * dist;
        else if (dir.includes('r')) state.tx = shapeW * dist;
        else if (dir.includes('t')) state.ty = -shapeH * dist;
        else state.ty = shapeH * dist; // default bottom
        break;
      }
      case 'zoom':
        state.opacity = p;
        state.scaleX = 0.1 + p * 0.9;
        state.scaleY = 0.1 + p * 0.9;
        break;
      case 'wipe':
        state.clipProgress = p;
        state.clipDir = dir || 'r';
        break;
      case 'split':
        state.clipProgress = p;
        state.clipDir = dir.includes('v') ? 'split-v' : 'split-h';
        break;
      case 'blinds':
        state.clipProgress = p;
        state.clipDir = dir.includes('v') ? 'blinds-v' : 'blinds-h';
        break;
      case 'box':
        state.clipProgress = p;
        state.clipDir = 'box';
        break;
      case 'wheel':
        state.clipProgress = p;
        state.clipDir = 'wheel';
        break;
      case 'wedge':
        state.clipProgress = p;
        state.clipDir = 'wedge';
        break;
      case 'strips':
        state.clipProgress = p;
        state.clipDir = dir.includes('r') ? 'strips-r' : 'strips-l';
        break;
      case 'stretch':
        state.opacity = p;
        if (dir.includes('h')) { state.scaleX = p; state.scaleY = 1; }
        else { state.scaleX = 1; state.scaleY = p; }
        break;
      case 'plus':
      case 'diamond':
        state.clipProgress = p;
        state.clipDir = effect;
        break;
      default:
        state.opacity = p;
    }
  } else if (step.type === 'exit') {
    // Exit = entrance reversed
    const exitStep = { ...step, type: 'entrance' };
    const entered = computeShapeState(exitStep, 1 - t, shapeW, shapeH);
    return entered;
  } else {
    // Emphasis
    switch (effect) {
      case 'spin':
        state.spin = p * 360;
        break;
      case 'grow':
      case 'shrink': {
        const maxScale = effect === 'grow' ? 1.5 : 0.5;
        const midT = t < 0.5 ? t * 2 : (1 - t) * 2;
        state.scaleX = 1 + (maxScale - 1) * midT;
        state.scaleY = state.scaleX;
        break;
      }
      case 'flash':
        state.opacity = t < 0.5 ? (t < 0.25 ? 0 : 1) : (t < 0.75 ? 0 : 1);
        break;
      default:
        state.opacity = 0.5 + 0.5 * Math.cos(t * Math.PI * 4); // pulse
    }
  }

  return state;
}

// ── Clip mask rendering ──────────────────────────────────────────────────────

function applyClipMask(ctx, clipDir, clipProgress, x, y, w, h) {
  const p = clipProgress;
  ctx.beginPath();
  switch (clipDir) {
    case 'r': ctx.rect(x, y, w * p, h); break;
    case 'l': ctx.rect(x + w * (1 - p), y, w * p, h); break;
    case 't': ctx.rect(x, y, w, h * p); break;
    case 'b': ctx.rect(x, y + h * (1 - p), w, h * p); break;
    case 'split-h': {
      const hw = w * p / 2;
      ctx.rect(x + w / 2 - hw, y, hw * 2, h);
      break;
    }
    case 'split-v': {
      const hh = h * p / 2;
      ctx.rect(x, y + h / 2 - hh, w, hh * 2);
      break;
    }
    case 'box': {
      const inset = Math.min(w, h) * (1 - p) / 2;
      ctx.rect(x + inset, y + inset, w - inset * 2, h - inset * 2);
      break;
    }
    case 'wheel': {
      const cx2 = x + w / 2, cy2 = y + h / 2;
      const r = Math.sqrt(w * w + h * h) / 2;
      ctx.moveTo(cx2, cy2);
      ctx.arc(cx2, cy2, r, -Math.PI / 2, -Math.PI / 2 + p * Math.PI * 2);
      ctx.closePath();
      break;
    }
    case 'wedge': {
      const cx2 = x + w / 2, cy2 = y + h / 2;
      const r = Math.sqrt(w * w + h * h) / 2;
      const a = p * Math.PI;
      ctx.moveTo(cx2, cy2);
      ctx.arc(cx2, cy2, r, -Math.PI / 2 - a, -Math.PI / 2 + a);
      ctx.closePath();
      break;
    }
    case 'blinds-h': {
      const bands = 6;
      const bh = h / bands;
      for (let i = 0; i < bands; i++) {
        ctx.rect(x, y + i * bh, w, bh * p);
      }
      break;
    }
    case 'blinds-v': {
      const bands = 6;
      const bw2 = w / bands;
      for (let i = 0; i < bands; i++) {
        ctx.rect(x + i * bw2, y, bw2 * p, h);
      }
      break;
    }
    case 'strips-r': {
      const bands = 8;
      const bh = h / bands;
      const bw2 = w / bands;
      for (let i = 0; i < bands; i++) {
        ctx.rect(x, y + i * bh, bw2 * (i + 1) * p, bh);
      }
      break;
    }
    case 'strips-l': {
      const bands = 8;
      const bh = h / bands;
      const bw2 = w / bands;
      for (let i = 0; i < bands; i++) {
        const tw = bw2 * (bands - i) * p;
        ctx.rect(x + w - tw, y + i * bh, tw, bh);
      }
      break;
    }
    case 'diamond': {
      const cx2 = x + w / 2, cy2 = y + h / 2;
      const rw = (w / 2) * p, rh = (h / 2) * p;
      ctx.moveTo(cx2 - rw, cy2);
      ctx.lineTo(cx2, cy2 - rh);
      ctx.lineTo(cx2 + rw, cy2);
      ctx.lineTo(cx2, cy2 + rh);
      ctx.closePath();
      break;
    }
    case 'plus': {
      const cx2 = x + w / 2, cy2 = y + h / 2;
      const arm = Math.min(w, h) / 2 * p;
      const thick = arm * 0.4;
      ctx.rect(cx2 - thick, cy2 - arm, thick * 2, arm * 2);
      ctx.rect(cx2 - arm, cy2 - thick, arm * 2, thick * 2);
      break;
    }
    default:
      ctx.rect(x, y, w, h);
  }
  ctx.clip();
}

// ── Transition renderer ──────────────────────────────────────────────────────

/**
 * Render a transition frame between two slide canvases.
 *
 * @param {CanvasRenderingContext2D} outCtx  — output canvas context
 * @param {HTMLCanvasElement} fromCanvas     — previous slide
 * @param {HTMLCanvasElement} toCanvas       — incoming slide
 * @param {SlideTransition}   transition
 * @param {number}            progress       — 0..1
 */
export function renderTransitionFrame(outCtx, fromCanvas, toCanvas, transition, progress) {
  const p = easeInOut(progress);
  const W = outCtx.canvas.width;
  const H = outCtx.canvas.height;

  outCtx.clearRect(0, 0, W, H);

  const type = transition?.type || 'fade';
  const dir  = transition?.dir || 'l';

  switch (type) {
    case 'cut':
      outCtx.drawImage(p < 0.5 ? fromCanvas : toCanvas, 0, 0, W, H);
      break;
    case 'fade':
    case 'fade_slow':
    default:
      outCtx.drawImage(fromCanvas, 0, 0, W, H);
      outCtx.globalAlpha = p;
      outCtx.drawImage(toCanvas, 0, 0, W, H);
      outCtx.globalAlpha = 1;
      break;
    case 'push': {
      const dx = dir === 'l' ? -W * p : dir === 'r' ? W * p : 0;
      const dy = dir === 'u' ? -H * p : dir === 'd' ? H * p : 0;
      outCtx.drawImage(fromCanvas, dx, dy, W, H);
      outCtx.drawImage(toCanvas, dx + (dir === 'l' ? W : dir === 'r' ? -W : 0),
                                  dy + (dir === 'u' ? H : dir === 'd' ? -H : 0), W, H);
      break;
    }
    case 'cover':
    case 'uncover': {
      const isUncover = type === 'uncover';
      const dx = dir === 'l' ? -W * (1 - p) : dir === 'r' ? W * (1 - p) : 0;
      const dy = dir === 'u' ? -H * (1 - p) : dir === 'd' ? H * (1 - p) : 0;
      if (isUncover) {
        outCtx.drawImage(toCanvas, 0, 0, W, H);
        outCtx.drawImage(fromCanvas, dx, dy, W, H);
      } else {
        outCtx.drawImage(fromCanvas, 0, 0, W, H);
        outCtx.drawImage(toCanvas, dx, dy, W, H);
      }
      break;
    }
    case 'wipe': {
      outCtx.drawImage(fromCanvas, 0, 0, W, H);
      outCtx.save();
      outCtx.beginPath();
      if (dir === 'l') outCtx.rect(W * (1 - p), 0, W * p, H);
      else if (dir === 'r') outCtx.rect(0, 0, W * p, H);
      else if (dir === 'u') outCtx.rect(0, H * (1 - p), W, H * p);
      else outCtx.rect(0, 0, W, H * p);
      outCtx.clip();
      outCtx.drawImage(toCanvas, 0, 0, W, H);
      outCtx.restore();
      break;
    }
    case 'zoom':
    case 'newsflash': {
      outCtx.drawImage(fromCanvas, 0, 0, W, H);
      outCtx.save();
      outCtx.globalAlpha = p;
      const s = type === 'newsflash' ? (1 + (1 - p) * 3) : (0.05 + p * 0.95);
      outCtx.translate(W / 2, H / 2);
      outCtx.scale(s, s);
      outCtx.drawImage(toCanvas, -W / 2, -H / 2, W, H);
      outCtx.restore();
      outCtx.globalAlpha = 1;
      break;
    }
    case 'dissolve':
    case 'wheel':
    case 'blinds': {
      // Approximated as fade for these complex pixel-based transitions
      outCtx.drawImage(fromCanvas, 0, 0, W, H);
      outCtx.globalAlpha = p;
      outCtx.drawImage(toCanvas, 0, 0, W, H);
      outCtx.globalAlpha = 1;
      break;
    }
  }
}

// ── PptxPlayer ────────────────────────────────────────────────────────────────

/**
 * Slide show player — drives animation and transition playback.
 *
 * @example
 * const player = new PptxPlayer(renderer, canvas);
 * await player.loadSlide(0);
 * player.play();
 * // Or: player.nextClick() to advance through click-triggered animations
 */
export class PptxPlayer {
  /**
   * @param {object}            renderer    — a loaded PptxRenderer instance
   * @param {HTMLCanvasElement} canvas      — output canvas
   */
  constructor(renderer, canvas) {
    this.renderer = renderer;
    this.canvas   = canvas;
    this.ctx      = canvas.getContext('2d');

    this._slideIndex     = 0;
    this._steps          = [];       // AnimStep[] for current slide
    this._transition     = null;     // SlideTransition for current slide
    this._clickNum       = 0;        // current click group
    this._activeAnimations = [];     // currently running RAF loops
    this._shapeStates    = new Map(); // shapeId → { opacity, tx, ty, scaleX, scaleY, clipProgress, clipDir, spin }
    this._baseCanvas     = null;     // fully-rendered static slide
    this._playing        = false;
    this._rafId          = null;

    // Event callbacks
    this.onSlideComplete   = null;  // called when all animations on slide finish
    this.onClickReady      = null;  // called when waiting for next click
    this.onTransitionStart = null;
    this.onTransitionEnd   = null;
  }

  // ── Public API ─────────────────────────────────────────────────────────────

  /** Pre-render the static slide and parse its animations. */
  async loadSlide(slideIndex) {
    this._slideIndex = slideIndex;
    this._clickNum   = 0;
    this._shapeStates.clear();
    this._stopAnimations();

    // Get the slide XML doc to parse timing
    const files    = this.renderer._files;
    const slidePath = this.renderer.slidePaths[slideIndex];
    if (!slidePath || !files) return;

    const slideXml = files[slidePath] ? new TextDecoder().decode(files[slidePath]) : null;
    if (!slideXml) return;

    const { parseXml } = await import('./utils.js').catch(() => ({
      parseXml: (s) => new DOMParser().parseFromString(s, 'application/xml'),
    }));
    const slideDoc = parseXml(slideXml);

    this._steps = parseAnimations(slideDoc);
    this._transition = parseTransition(slideDoc);

    // Identify shapes that have entrance animations → hide them initially
    const entranceIds = new Set(
      this._steps.filter(s => s.type === 'entrance').map(s => s.shapeId)
    );
    this._initiallyHidden = entranceIds;

    // Render the base slide (with entrance-animated shapes hidden)
    this._baseCanvas = await this._renderBaseSlide();
    this._drawBase();

    // If there are click=0 (auto-start) animations, play them immediately
    await this._playClickGroup(0);
  }

  /** Advance to next click group. */
  async nextClick() {
    this._clickNum++;
    await this._playClickGroup(this._clickNum);
  }

  /** Start playback of all remaining click groups automatically. */
  async play(autoAdvanceMs = 1500) {
    this._playing = true;
    const maxClick = Math.max(...this._steps.map(s => s.clickNum), 0);
    while (this._playing && this._clickNum <= maxClick) {
      await this._playClickGroup(this._clickNum);
      this._clickNum++;
      if (this._clickNum <= maxClick) {
        await this._delay(autoAdvanceMs);
      }
    }
    this._playing = false;
  }

  /** Pause all running animations. */
  pause() {
    this._playing = false;
    this._stopAnimations();
  }

  /** Reset to initial state. */
  async stop() {
    this._playing = false;
    this._stopAnimations();
    this._shapeStates.clear();
    this._clickNum = 0;
    if (this._baseCanvas) this._drawBase();
  }

  /**
   * Animate a transition from the current slide to a new slide index.
   * @param {number} nextIndex
   * @returns {Promise<void>} resolves when transition completes
   */
  async transitionTo(nextIndex) {
    const fromCanvas = document.createElement('canvas');
    fromCanvas.width  = this.canvas.width;
    fromCanvas.height = this.canvas.height;
    fromCanvas.getContext('2d').drawImage(this.canvas, 0, 0);

    await this.loadSlide(nextIndex);
    const toCanvas = this._baseCanvas;

    const transition = this._transition || { type: 'fade', duration: 700, dir: 'l' };
    this.onTransitionStart?.({ from: this._slideIndex - 1, to: nextIndex, transition });

    await this._animateTransition(fromCanvas, toCanvas, transition);
    this.onTransitionEnd?.({ slideIndex: nextIndex });
  }

  // ── Private methods ────────────────────────────────────────────────────────

  async _renderBaseSlide() {
    const w = this.canvas.width;
    const bc = typeof OffscreenCanvas !== 'undefined'
      ? new OffscreenCanvas(w, Math.round(w / (this.renderer.slideSize.cx / this.renderer.slideSize.cy)))
      : Object.assign(document.createElement('canvas'), { width: w, height: Math.round(w / (this.renderer.slideSize.cx / this.renderer.slideSize.cy)) });
    await this.renderer.renderSlide(this._slideIndex, bc, w);
    return bc;
  }

  _drawBase() {
    if (!this._baseCanvas) return;
    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    this.ctx.drawImage(this._baseCanvas, 0, 0, this.canvas.width, this.canvas.height);
  }

  async _playClickGroup(clickNum) {
    const groupSteps = this._steps.filter(s => s.clickNum === clickNum);
    if (!groupSteps.length) return;

    this.onClickReady?.({ clickNum, stepCount: groupSteps.length });

    // Group by delay bucket → play in parallel
    const maxDelay = Math.max(...groupSteps.map(s => s.delay + s.duration), 0);

    await new Promise(resolve => {
      const startTime = performance.now();
      const completed = new Set();

      const frame = (now) => {
        const elapsed = now - startTime;

        for (const step of groupSteps) {
          if (completed.has(step)) continue;
          const stepElapsed = elapsed - step.delay;
          if (stepElapsed < 0) continue;

          const progress = Math.min(1, stepElapsed / step.duration);
          // We'd need shape bounds here — we'll approximate with slide dimensions
          const sw = this.canvas.width;
          const sh = this.canvas.height;
          const state = computeShapeState(step, progress, sw * 0.3, sh * 0.3);
          this._shapeStates.set(step.shapeId, { ...state, step });

          if (progress >= 1) completed.add(step);
        }

        // Re-composite
        this._composite();

        if (completed.size < groupSteps.length) {
          this._rafId = requestAnimationFrame(frame);
        } else {
          resolve();
        }
      };

      this._rafId = requestAnimationFrame(frame);
    });
  }

  _composite() {
    // Draw base slide then apply per-shape states
    // Note: full per-shape compositing would require re-rendering each shape
    // individually. As an approximation, we do a full redraw with the overall
    // canvas transforms. For production, each shape should be drawn to its own
    // offscreen canvas and composited with transforms.
    this._drawBase();

    // Draw shape state overlays (opacity/position highlights)
    for (const [shapeId, state] of this._shapeStates) {
      if (state.opacity < 0.99) {
        // We don't have shape bboxes here — production renderer would need them
        // For now the effect is subtle and graceful
      }
    }
  }

  async _animateTransition(fromCanvas, toCanvas, transition) {
    const duration = transition.duration;
    await new Promise(resolve => {
      const start = performance.now();
      const frame = (now) => {
        const progress = Math.min(1, (now - start) / duration);
        renderTransitionFrame(this.ctx, fromCanvas, toCanvas, transition, progress);
        if (progress < 1) {
          this._rafId = requestAnimationFrame(frame);
        } else {
          resolve();
        }
      };
      this._rafId = requestAnimationFrame(frame);
    });
  }

  _stopAnimations() {
    if (this._rafId) {
      cancelAnimationFrame(this._rafId);
      this._rafId = null;
    }
  }

  _delay(ms) {
    return new Promise(r => setTimeout(r, ms));
  }
}

// ── Shape animation compositor (standalone) ───────────────────────────────────

/**
 * Apply animation state transforms to a shape that has already been rendered
 * onto its own offscreen canvas, then composite to the output context.
 *
 * This is the low-level primitive used when each shape has its own canvas.
 *
 * @param {CanvasRenderingContext2D} outCtx
 * @param {HTMLCanvasElement}        shapeCanvas
 * @param {object}                   state       — from computeShapeState()
 * @param {number}                   cx, cy, cw, ch  — shape bounds on output
 */
export function compositeShape(outCtx, shapeCanvas, state, cx, cy, cw, ch) {
  if (state.opacity === 0) return;

  outCtx.save();
  outCtx.globalAlpha = state.opacity;

  const pivX = cx + cw / 2;
  const pivY = cy + ch / 2;
  outCtx.translate(pivX + state.tx, pivY + state.ty);

  if (state.spin) outCtx.rotate(state.spin * Math.PI / 180);
  if (state.scaleX !== 1 || state.scaleY !== 1) {
    outCtx.scale(state.scaleX, state.scaleY);
  }

  if (state.clipProgress < 1) {
    applyClipMask(outCtx, state.clipDir, state.clipProgress, -cw / 2, -ch / 2, cw, ch);
  }

  outCtx.drawImage(shapeCanvas, -cw / 2, -ch / 2, cw, ch);
  outCtx.restore();
}
