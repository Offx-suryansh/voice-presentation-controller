/**
 * slideshow.js — Full-screen slide show player.
 *
 * Creates a browser-native presentation mode with:
 *   - Fullscreen API integration (F11 / Escape to exit)
 *   - Keyboard navigation (←→ arrow keys, Space, PageUp/Down, Home, End)
 *   - Touch / swipe support
 *   - Slide counter HUD
 *   - Presenter notes panel (optional)
 *   - Animation engine integration (PptxPlayer)
 *   - Transition effects between slides
 *   - Thumbnail strip for quick navigation
 *
 * Usage:
 *   import { SlideShow } from 'pptx-canvas-renderer';
 *
 *   const show = new SlideShow(renderer, container);
 *   await show.start(0);     // start from slide 0, request fullscreen
 *   show.stop();             // exit and clean up
 *
 *   // Or without fullscreen:
 *   const show = new SlideShow(renderer, container, { fullscreen: false });
 *   await show.start();
 */

export class SlideShow {
  /**
   * @param {object}          renderer  — loaded PptxRenderer instance
   * @param {HTMLElement}     container — DOM element to attach to
   * @param {object}          [opts]
   * @param {boolean}         [opts.fullscreen=true]    — request fullscreen on start
   * @param {boolean}         [opts.showNotes=false]    — show presenter notes panel
   * @param {boolean}         [opts.showThumbs=false]   — show thumbnail strip
   * @param {boolean}         [opts.showHud=true]       — show slide counter HUD
   * @param {boolean}         [opts.loop=false]         — loop back to start at end
   * @param {boolean}         [opts.autoAdvance=0]      — ms between slides (0=manual)
   * @param {function}        [opts.onSlideChange]      — (index) => void
   */
  constructor(renderer, container, opts = {}) {
    this.renderer  = renderer;
    this.container = container;
    this.opts = {
      fullscreen:  true,
      showNotes:   false,
      showThumbs:  false,
      showHud:     true,
      loop:        false,
      autoAdvance: 0,
      ...opts,
    };

    this._index       = 0;
    this._playing     = false;
    this._player      = null;
    this._autoTimer   = null;
    this._el          = null; // root overlay element
    this._canvas      = null;
    this._notesEl     = null;
    this._hudEl       = null;
    this._thumbsEl    = null;
    this._thumbnails  = []; // low-res canvas elements

    // Touch tracking
    this._touchStartX = 0;
    this._touchStartY = 0;

    // Bound handlers (for removeEventListener)
    this._onKey       = this._onKey.bind(this);
    this._onResize    = this._onResize.bind(this);
    this._onFsChange  = this._onFsChange.bind(this);
    this._onTouchStart = this._onTouchStart.bind(this);
    this._onTouchEnd   = this._onTouchEnd.bind(this);
  }

  // ── Lifecycle ───────────────────────────────────────────────────────────────

  /** Start the slide show, optionally requesting fullscreen. */
  async start(slideIndex = 0) {
    this._playing = true;
    this._index   = Math.max(0, Math.min(slideIndex, this.renderer.slideCount - 1));

    this._buildDOM();
    this._attachEvents();

    if (this.opts.fullscreen) {
      await this._requestFullscreen();
    }

    // Pre-generate thumbnails in background
    if (this.opts.showThumbs) {
      this._generateThumbnails(); // don't await — background task
    }

    await this._renderCurrent();
    this._updateHud();

    if (this.opts.autoAdvance > 0) {
      this._startAutoAdvance();
    }
  }

  /** Stop and clean up. */
  stop() {
    this._playing = false;
    this._stopAutoAdvance();
    this._detachEvents();
    if (document.fullscreenElement === this._el) {
      document.exitFullscreen().catch(() => {});
    }
    if (this._el && this._el.parentNode) {
      this._el.parentNode.removeChild(this._el);
    }
    this._el     = null;
    this._canvas = null;
    this._player = null;
    this.opts.onSlideChange?.(null);
  }

  /** Go to a specific slide. */
  async goto(index) {
    if (!this._playing) return;
    const newIndex = Math.max(0, Math.min(index, this.renderer.slideCount - 1));
    if (newIndex === this._index) return;

    const prevIndex = this._index;
    this._index = newIndex;
    this._resetAutoAdvance();
    await this._renderCurrent(prevIndex);
    this._updateHud();
    this.opts.onSlideChange?.(this._index);
  }

  /** Advance to next slide (or next animation click). */
  async next() {
    if (!this._playing) return;
    if (this._player) {
      // If animations remain, advance click
      const steps = this.renderer.getAnimations?.(this._index) || [];
      const maxClick = steps.length ? Math.max(...steps.map(s => s.clickNum), 0) : 0;
      // Simple heuristic — if player exists and not at end, click
      await this._player.nextClick?.();
    }
    if (this._index < this.renderer.slideCount - 1) {
      await this.goto(this._index + 1);
    } else if (this.opts.loop) {
      await this.goto(0);
    }
  }

  /** Go to previous slide. */
  async prev() {
    if (!this._playing) return;
    if (this._index > 0) {
      await this.goto(this._index - 1);
    } else if (this.opts.loop) {
      await this.goto(this.renderer.slideCount - 1);
    }
  }

  get currentIndex() { return this._index; }
  get isPlaying()    { return this._playing; }

  // ── DOM construction ────────────────────────────────────────────────────────

  _buildDOM() {
    // Root overlay
    const el = document.createElement('div');
    el.style.cssText = [
      'position:fixed', 'inset:0', 'z-index:999999',
      'background:#000', 'display:flex', 'flex-direction:column',
      'align-items:center', 'justify-content:center',
      'user-select:none', 'touch-action:none',
    ].join(';');
    el.setAttribute('tabindex', '0');
    this._el = el;

    // Main canvas area
    const canvasWrap = document.createElement('div');
    canvasWrap.style.cssText = 'position:relative; flex:1; display:flex; align-items:center; justify-content:center; width:100%; overflow:hidden;';

    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'display:block; box-shadow:0 4px 32px rgba(0,0,0,0.6);';
    this._canvas = canvas;
    canvasWrap.appendChild(canvas);
    el.appendChild(canvasWrap);

    // HUD (slide counter + controls)
    if (this.opts.showHud) {
      const hud = document.createElement('div');
      hud.style.cssText = [
        'position:absolute', 'bottom:24px', 'left:50%', 'transform:translateX(-50%)',
        'display:flex', 'align-items:center', 'gap:12px',
        'background:rgba(0,0,0,0.55)', 'backdrop-filter:blur(8px)',
        'border-radius:24px', 'padding:8px 20px',
        'color:white', 'font:500 14px/1 system-ui,sans-serif',
        'pointer-events:none',
      ].join(';');
      this._hudEl = hud;
      el.appendChild(hud);
    }

    // Notes panel
    if (this.opts.showNotes) {
      const notes = document.createElement('div');
      notes.style.cssText = [
        'width:100%', 'max-height:22vh', 'overflow-y:auto',
        'background:rgba(0,0,0,0.7)', 'backdrop-filter:blur(8px)',
        'color:#e0e0e0', 'font:14px/1.5 system-ui,sans-serif',
        'padding:12px 24px', 'white-space:pre-wrap', 'flex-shrink:0',
      ].join(';');
      this._notesEl = notes;
      el.appendChild(notes);
    }

    // Thumbnail strip
    if (this.opts.showThumbs) {
      const thumbs = document.createElement('div');
      thumbs.style.cssText = [
        'display:flex', 'gap:6px', 'padding:8px 16px',
        'overflow-x:auto', 'background:rgba(0,0,0,0.7)',
        'width:100%', 'flex-shrink:0',
      ].join(';');
      this._thumbsEl = thumbs;
      el.appendChild(thumbs);
      // Placeholder thumbnails
      for (let i = 0; i < this.renderer.slideCount; i++) {
        const thumb = document.createElement('canvas');
        thumb.width  = 120;
        thumb.height = Math.round(120 / (this.renderer.slideSize.cx / this.renderer.slideSize.cy));
        thumb.style.cssText = 'flex-shrink:0; cursor:pointer; border:2px solid transparent; border-radius:3px; opacity:0.6; transition:opacity 0.2s,border-color 0.2s;';
        thumb.title = `Slide ${i + 1}`;
        const idx = i;
        thumb.addEventListener('click', () => this.goto(idx));
        thumbs.appendChild(thumb);
        this._thumbnails.push(thumb);
      }
    }

    // Click to advance (on canvas wrap)
    canvasWrap.addEventListener('click', () => this.next());

    // Nav arrow buttons (hidden unless hover)
    this._buildNavArrows(canvasWrap);

    // Close button
    const closeBtn = document.createElement('button');
    closeBtn.textContent = '✕';
    closeBtn.style.cssText = [
      'position:absolute', 'top:16px', 'right:20px',
      'background:rgba(255,255,255,0.15)', 'border:none',
      'color:white', 'font-size:18px', 'width:36px', 'height:36px',
      'border-radius:50%', 'cursor:pointer', 'opacity:0.7',
      'transition:opacity 0.2s',
    ].join(';');
    closeBtn.addEventListener('mouseover', () => closeBtn.style.opacity = '1');
    closeBtn.addEventListener('mouseout',  () => closeBtn.style.opacity = '0.7');
    closeBtn.addEventListener('click', (e) => { e.stopPropagation(); this.stop(); });
    el.appendChild(closeBtn);

    this.container.appendChild(el);
    el.focus();
    this._resizeCanvas();
  }

  _buildNavArrows(parent) {
    const makeArrow = (dir) => {
      const btn = document.createElement('button');
      btn.textContent = dir === 'prev' ? '❮' : '❯';
      btn.style.cssText = [
        'position:absolute',
        dir === 'prev' ? 'left:16px' : 'right:16px',
        'top:50%', 'transform:translateY(-50%)',
        'background:rgba(255,255,255,0.18)', 'border:none',
        'color:white', 'font-size:22px', 'width:48px', 'height:64px',
        'border-radius:8px', 'cursor:pointer',
        'opacity:0', 'transition:opacity 0.2s',
      ].join(';');
      btn.addEventListener('mouseenter', () => btn.style.opacity = '1');
      btn.addEventListener('mouseleave', () => btn.style.opacity = '0');
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        dir === 'prev' ? this.prev() : this.next();
      });
      parent.addEventListener('mouseenter', () => btn.style.opacity = '0.5');
      parent.addEventListener('mouseleave', () => btn.style.opacity = '0');
      parent.appendChild(btn);
    };
    makeArrow('prev');
    makeArrow('next');
  }

  // ── Canvas sizing ───────────────────────────────────────────────────────────

  _resizeCanvas() {
    if (!this._canvas || !this._el) return;
    const { cx, cy } = this.renderer.slideSize;
    const aspect = cx / cy;

    const container = this._canvas.parentElement;
    const availW = container.clientWidth   || window.innerWidth;
    const availH = container.clientHeight  || window.innerHeight * 0.75;

    let w = availW, h = w / aspect;
    if (h > availH) { h = availH; w = h * aspect; }

    this._canvas.width  = Math.round(w * window.devicePixelRatio);
    this._canvas.height = Math.round(h * window.devicePixelRatio);
    this._canvas.style.width  = Math.round(w) + 'px';
    this._canvas.style.height = Math.round(h) + 'px';
  }

  // ── Rendering ───────────────────────────────────────────────────────────────

  async _renderCurrent(prevIndex = null) {
    if (!this._canvas) return;
    this._resizeCanvas();

    // Transition if we have a prev canvas snapshot
    if (prevIndex !== null && this.renderer.getTransition) {
      const transition = this.renderer.getTransition(this._index);
      if (transition && transition.type !== 'cut' && typeof OffscreenCanvas !== 'undefined') {
        const { renderTransitionFrame } = await import('./animation.js');
        const from = new OffscreenCanvas(this._canvas.width, this._canvas.height);
        from.getContext('2d').drawImage(this._canvas, 0, 0);

        await this.renderer.renderSlide(this._index, this._canvas, this._canvas.width / window.devicePixelRatio);

        const to = new OffscreenCanvas(this._canvas.width, this._canvas.height);
        to.getContext('2d').drawImage(this._canvas, 0, 0);

        const ctx = this._canvas.getContext('2d');
        const dur = transition.duration || 700;
        const start = performance.now();
        await new Promise(resolve => {
          const frame = (now) => {
            const p = Math.min(1, (now - start) / dur);
            renderTransitionFrame(ctx, from, to, transition, p);
            if (p < 1) requestAnimationFrame(frame); else resolve();
          };
          requestAnimationFrame(frame);
        });
        this._updateThumbnail(this._index);
        this._updateNotes();
        return;
      }
    }

    await this.renderer.renderSlide(this._index, this._canvas, this._canvas.width / window.devicePixelRatio);
    this._updateThumbnail(this._index);
    this._updateNotes();
  }

  _updateThumbnail(index) {
    const thumb = this._thumbnails[index];
    if (!thumb || !this._canvas) return;
    thumb.getContext('2d').drawImage(this._canvas, 0, 0, thumb.width, thumb.height);
    // Update active state
    this._thumbnails.forEach((t, i) => {
      t.style.borderColor  = i === this._index ? '#4af' : 'transparent';
      t.style.opacity      = i === this._index ? '1' : '0.6';
    });
  }

  _updateHud() {
    if (!this._hudEl) return;
    const n = this.renderer.slideCount;
    this._hudEl.innerHTML = `
      <span style="opacity:0.7">Slide</span>
      <strong>${this._index + 1}</strong>
      <span style="opacity:0.5">of ${n}</span>
    `;
  }

  async _updateNotes() {
    if (!this._notesEl) return;
    try {
      const notes = await this.renderer.getSlideNotes(this._index);
      this._notesEl.textContent = notes || '(no notes)';
    } catch (_) {
      this._notesEl.textContent = '';
    }
  }

  async _generateThumbnails() {
    for (let i = 0; i < this.renderer.slideCount; i++) {
      if (!this._playing) break;
      const thumb = this._thumbnails[i];
      if (!thumb) continue;
      try {
        await this.renderer.renderSlide(i, thumb, thumb.width);
      } catch (_) {}
      // Small delay to avoid blocking the main render
      await new Promise(r => setTimeout(r, 50));
    }
  }

  // ── Keyboard & touch ────────────────────────────────────────────────────────

  _attachEvents() {
    document.addEventListener('keydown', this._onKey);
    window.addEventListener('resize', this._onResize);
    document.addEventListener('fullscreenchange', this._onFsChange);
    if (this._el) {
      this._el.addEventListener('touchstart', this._onTouchStart, { passive: true });
      this._el.addEventListener('touchend',   this._onTouchEnd,   { passive: true });
    }
  }

  _detachEvents() {
    document.removeEventListener('keydown', this._onKey);
    window.removeEventListener('resize', this._onResize);
    document.removeEventListener('fullscreenchange', this._onFsChange);
  }

  _onKey(e) {
    if (!this._playing) return;
    switch (e.key) {
      case 'ArrowRight': case 'ArrowDown': case ' ': case 'PageDown':
        e.preventDefault(); this.next(); break;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace':
        e.preventDefault(); this.prev(); break;
      case 'Home':  e.preventDefault(); this.goto(0); break;
      case 'End':   e.preventDefault(); this.goto(this.renderer.slideCount - 1); break;
      case 'Escape': case 'q': case 'Q':
        e.preventDefault(); this.stop(); break;
      case 'f': case 'F':
        e.preventDefault(); this._toggleFullscreen(); break;
    }
  }

  _onResize() {
    this._resizeCanvas();
    this._renderCurrent(); // re-render at new size
  }

  _onFsChange() {
    if (!document.fullscreenElement) {
      // User pressed Esc in fullscreen — stop show
      this.stop();
    }
  }

  _onTouchStart(e) {
    this._touchStartX = e.touches[0].clientX;
    this._touchStartY = e.touches[0].clientY;
  }

  _onTouchEnd(e) {
    const dx = e.changedTouches[0].clientX - this._touchStartX;
    const dy = e.changedTouches[0].clientY - this._touchStartY;
    if (Math.abs(dx) > 40 && Math.abs(dx) > Math.abs(dy)) {
      if (dx < 0) this.next(); else this.prev();
    } else if (Math.abs(dx) < 10 && Math.abs(dy) < 10) {
      this.next(); // tap to advance
    }
  }

  // ── Fullscreen ──────────────────────────────────────────────────────────────

  async _requestFullscreen() {
    try {
      const el = this._el;
      if (el.requestFullscreen)            await el.requestFullscreen();
      else if (el.webkitRequestFullscreen) await el.webkitRequestFullscreen();
    } catch (_) { /* ignore if not allowed */ }
  }

  async _toggleFullscreen() {
    if (document.fullscreenElement) {
      await document.exitFullscreen().catch(() => {});
    } else {
      await this._requestFullscreen();
    }
  }

  // ── Auto-advance ─────────────────────────────────────────────────────────────

  _startAutoAdvance() {
    this._stopAutoAdvance();
    if (this.opts.autoAdvance > 0) {
      this._autoTimer = setInterval(() => this.next(), this.opts.autoAdvance);
    }
  }

  _stopAutoAdvance() {
    if (this._autoTimer) { clearInterval(this._autoTimer); this._autoTimer = null; }
  }

  _resetAutoAdvance() {
    if (this.opts.autoAdvance > 0) this._startAutoAdvance();
  }
}
