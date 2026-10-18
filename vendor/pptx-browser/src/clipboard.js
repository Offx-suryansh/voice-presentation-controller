/**
 * clipboard.js — Clipboard copy and progressive lazy rendering.
 *
 * CLIPBOARD API
 * ─────────────
 * copySlideToClipboard(slideIndex, renderer)
 *   Renders the slide and writes it as a PNG image to the system clipboard.
 *   Requires Clipboard API permission (shown to user by browser).
 *
 * copyAllSlidesToClipboard(renderer)
 *   Copies all slides as a ZIP of PNGs to clipboard text (data URL list).
 *   (Clipboard doesn't support multi-image, so this creates a JSON manifest.)
 *
 * PROGRESSIVE / LAZY RENDERING
 * ─────────────────────────────
 * createLazyDeck(renderer, container, opts)
 *   Builds a scrollable deck view where slides are rendered on-demand as
 *   they scroll into the viewport, using IntersectionObserver.
 *   Shows a low-res placeholder until the full render is ready.
 *
 * Usage:
 *   // Copy slide 3 to clipboard
 *   await copySlideToClipboard(2, renderer);
 *
 *   // Progressive deck view (great for large PPTX files)
 *   const deck = createLazyDeck(renderer, document.getElementById('deck'));
 *   // Each slide renders as you scroll to it
 *   deck.destroy(); // clean up
 *
 *   // Or via renderer instance:
 *   await renderer.copySlide(2);
 *   const deck = renderer.createDeck(container);
 */

// ── DPI helper ──────────────────────────────────────────────────────────────────

function dpiToWidth(renderer, dpi) {
  const inches = renderer.slideSize.cx / 914400;
  return Math.round(inches * dpi);
}

// ── Clipboard ─────────────────────────────────────────────────────────────────

/**
 * Copy a single slide to the clipboard as a PNG image.
 *
 * Requires Clipboard API and clipboard-write permission.
 * Falls back to opening the image in a new tab if clipboard is unavailable.
 *
 * @param {number} slideIndex
 * @param {object} renderer
 * @param {object|number} [opts]  options or legacy pixel width
 * @param {number} [opts.width]    pixel width (overrides dpi)
 * @param {number} [opts.dpi=150]  dots per inch
 * @returns {Promise<{success: boolean, method: string}>}
 */
export async function copySlideToClipboard(slideIndex, renderer, opts = {}) {
  // Legacy: allow passing a raw number as width
  if (typeof opts === 'number') opts = { width: opts };
  const { width = null, dpi = 150 } = opts;
  const resolvedWidth = width ?? dpiToWidth(renderer, dpi);
  // Render to an offscreen canvas
  const canvas = await _renderToCanvas(slideIndex, renderer, resolvedWidth);
  const dataUrl = canvas.toDataURL('image/png');

  // Method 1: Clipboard API (modern browsers, requires permission)
  if (navigator.clipboard && window.ClipboardItem) {
    try {
      const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
      await navigator.clipboard.write([
        new ClipboardItem({ 'image/png': blob }),
      ]);
      return { success: true, method: 'clipboard-api', dataUrl };
    } catch (err) {
      if (err.name !== 'NotAllowedError') throw err;
      // Permission denied — fall through to fallback
    }
  }

  // Method 2: execCommand fallback (deprecated but still works in some contexts)
  if (document.execCommand) {
    try {
      const img = new Image();
      img.src = dataUrl;
      await new Promise(r => { img.onload = r; });
      // Create a contenteditable div, paste the image, copy it
      const div = document.createElement('div');
      div.contentEditable = 'true';
      div.style.cssText = 'position:fixed;left:-9999px;top:-9999px;opacity:0;';
      div.appendChild(img);
      document.body.appendChild(div);
      const range = document.createRange();
      range.selectNodeContents(div);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      const ok = document.execCommand('copy');
      document.body.removeChild(div);
      sel.removeAllRanges();
      if (ok) return { success: true, method: 'execCommand', dataUrl };
    } catch (_) {}
  }

  // Method 3: Open in new tab (always works)
  const win = window.open();
  if (win) {
    win.document.write(`<img src="${dataUrl}" style="max-width:100%">`);
    win.document.title = `Slide ${slideIndex + 1}`;
  }
  return { success: false, method: 'opened-tab', dataUrl };
}

/**
 * Download a slide as a PNG file.
 *
 * @param {number} slideIndex
 * @param {object} renderer
 * @param {object|number} [opts]  options or legacy pixel width
 * @param {number} [opts.width]    pixel width (overrides dpi)
 * @param {number} [opts.dpi=300]  dots per inch (default 300 for download)
 * @param {string} [opts.filename]
 */
export async function downloadSlide(slideIndex, renderer, opts = {}, filename) {
  if (typeof opts === 'number') opts = { width: opts };
  const { width = null, dpi = 300, filename: fn } = opts;
  const resolvedWidth = width ?? dpiToWidth(renderer, dpi);
  const canvas = await _renderToCanvas(slideIndex, renderer, resolvedWidth);
  const dataUrl = canvas.toDataURL('image/png');
  const name = filename || fn || `slide-${slideIndex + 1}.png`;
  const a = document.createElement('a');
  a.href = dataUrl;
  a.download = name;
  a.click();
}

/**
 * Download all slides as individual PNG files, or as a ZIP if JSZip is available.
 *
 * @param {object} renderer
 * @param {object|number} [opts]  options or legacy pixel width
 * @param {number} [opts.width]    pixel width (overrides dpi)
 * @param {number} [opts.dpi=300]  dots per inch
 * @param {function} [opts.onProgress] (completed, total) => void
 */
export async function downloadAllSlides(renderer, opts = {}, onProgress) {
  if (typeof opts === 'number') opts = { width: opts };
  const { onProgress: progFn, ...slideOpts } = opts;
  const progress = onProgress || progFn;
  const n = renderer.slideCount;
  for (let i = 0; i < n; i++) {
    await downloadSlide(i, renderer, slideOpts);
    progress?.(i + 1, n);
    // Small delay to allow browser to process download
    await new Promise(r => setTimeout(r, 100));
  }
}

/**
 * Render a slide to an HTMLCanvasElement.
 */
async function _renderToCanvas(slideIndex, renderer, width) {
  const { cx, cy } = renderer.slideSize;
  const aspect = cx / cy;
  const h = Math.round(width / aspect);

  const canvas = document.createElement('canvas');
  canvas.width  = width;
  canvas.height = h;
  await renderer.renderSlide(slideIndex, canvas, width);
  return canvas;
}

// ── Progressive lazy rendering ────────────────────────────────────────────────

/**
 * @typedef {object} LazyDeckOpts
 * @property {number}   [thumbWidth=320]    px width of initial low-res render
 * @property {number}   [fullWidth=1280]    px width of full-res render
 * @property {string}   [gap='24px']        CSS gap between slides
 * @property {string}   [background='#1a1a1a'] container background
 * @property {string}   [slideBackground='#fff'] per-slide background
 * @property {boolean}  [shadow=true]       drop shadow on slides
 * @property {boolean}  [clickToShow=false] clicking a slide opens fullscreen
 * @property {string}   [maxWidth='900px']  max width of slide display
 * @property {number}   [rootMargin=200]    px pre-load margin (IntersectionObserver)
 * @property {function} [onSlideVisible]    (index) => void
 * @property {function} [onSlideRendered]   (index) => void
 */

/**
 * Create a progressive lazy-rendered deck view.
 *
 * Slides are shown as grey placeholders and rendered on-demand as they
 * scroll into the viewport (plus a small lookahead margin).
 *
 * Returns a controller object with:
 *   .destroy()             — remove everything and clean up
 *   .scrollTo(index)       — smooth scroll to a slide
 *   .renderAll()           — force render all slides immediately
 *   .getCanvas(index)      — get the canvas element for a slide
 *
 * @param {object}      renderer
 * @param {HTMLElement} container
 * @param {LazyDeckOpts} [opts]
 * @returns {LazyDeckController}
 */
export function createLazyDeck(renderer, container, opts = {}) {
  const {
    thumbWidth    = 320,
    fullWidth     = 1280,
    gap           = '24px',
    background    = '#1a1a1a',
    slideBackground = '#fff',
    shadow        = true,
    clickToShow   = false,
    maxWidth      = '900px',
    rootMargin    = 200,
    onSlideVisible,
    onSlideRendered,
  } = opts;

  const { cx, cy } = renderer.slideSize;
  const aspect = cx / cy;
  const n = renderer.slideCount;

  // ── Build container ────────────────────────────────────────────────────────

  container.style.cssText = `
    background:${background};
    padding:${gap};
    display:flex;
    flex-direction:column;
    align-items:center;
    gap:${gap};
    overflow-y:auto;
    position:relative;
  `;

  // ── Slide elements ────────────────────────────────────────────────────────

  const slideWrappers = [];
  const canvases      = [];
  const renderStates  = []; // 'pending' | 'thumb' | 'full'

  for (let i = 0; i < n; i++) {
    const wrap = document.createElement('div');
    wrap.style.cssText = `
      width:100%;
      max-width:${maxWidth};
      position:relative;
      background:${slideBackground};
      border-radius:4px;
      ${shadow ? 'box-shadow:0 4px 24px rgba(0,0,0,0.5);' : ''}
      overflow:hidden;
      aspect-ratio:${cx}/${cy};
      flex-shrink:0;
    `;
    wrap.setAttribute('data-slide', i);

    // Slide number badge
    const badge = document.createElement('div');
    badge.style.cssText = `
      position:absolute;bottom:8px;right:10px;
      background:rgba(0,0,0,0.45);color:#fff;
      font:11px/1.4 system-ui,sans-serif;
      padding:2px 7px;border-radius:10px;
      pointer-events:none;z-index:1;
    `;
    badge.textContent = i + 1;
    wrap.appendChild(badge);

    // Canvas (sized to the wrapper)
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'display:block;width:100%;height:100%;';
    canvas.width  = thumbWidth;
    canvas.height = Math.round(thumbWidth / aspect);
    wrap.appendChild(canvas);

    // Placeholder while not yet rendered
    const placeholder = _buildPlaceholder(i, cx, cy, slideBackground);
    wrap.appendChild(placeholder);

    // Click handler
    if (clickToShow) {
      wrap.style.cursor = 'pointer';
      wrap.addEventListener('click', () => {
        const show = new (require('./slideshow.js').SlideShow)(renderer, document.body);
        show.start(i);
      });
    }

    container.appendChild(wrap);
    slideWrappers.push(wrap);
    canvases.push(canvas);
    renderStates.push('pending');
  }

  // ── IntersectionObserver ──────────────────────────────────────────────────

  const renderQueue = [];
  let renderBusy = false;

  async function processQueue() {
    if (renderBusy || renderQueue.length === 0) return;
    renderBusy = true;
    while (renderQueue.length > 0) {
      const idx = renderQueue.shift();
      if (renderStates[idx] === 'full') continue;

      const canvas = canvases[idx];
      const wrap   = slideWrappers[idx];
      const placeholder = wrap.querySelector('[data-placeholder]');

      try {
        // Render at full resolution
        canvas.width  = fullWidth;
        canvas.height = Math.round(fullWidth / aspect);
        await renderer.renderSlide(idx, canvas, fullWidth);
        renderStates[idx] = 'full';

        // Fade in
        canvas.style.transition = 'opacity 0.3s';
        canvas.style.opacity = '1';
        if (placeholder) placeholder.style.display = 'none';

        onSlideRendered?.(idx);
      } catch (err) {
        console.warn(`LazyDeck: failed to render slide ${idx}`, err);
      }
    }
    renderBusy = false;
  }

  const observer = new IntersectionObserver((entries) => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      const idx = parseInt(entry.target.getAttribute('data-slide'), 10);
      if (isNaN(idx) || renderStates[idx] === 'full') continue;

      onSlideVisible?.(idx);

      // Enqueue: current slide first, then neighbours
      for (const neighbor of [idx, idx - 1, idx + 1, idx + 2].filter(j => j >= 0 && j < n)) {
        if (renderStates[neighbor] !== 'full' && !renderQueue.includes(neighbor)) {
          renderQueue.push(neighbor);
        }
      }
      processQueue();
    }
  }, {
    root: container,
    rootMargin: `${rootMargin}px`,
    threshold: 0,
  });

  slideWrappers.forEach(w => observer.observe(w));

  // ── Controller ─────────────────────────────────────────────────────────────

  return {
    destroy() {
      observer.disconnect();
      container.innerHTML = '';
    },

    scrollTo(index, behavior = 'smooth') {
      const wrap = slideWrappers[index];
      if (wrap) wrap.scrollIntoView({ behavior, block: 'start' });
    },

    async renderAll(onProgress) {
      for (let i = 0; i < n; i++) {
        if (renderStates[i] !== 'full') {
          renderQueue.push(i);
        }
      }
      await processQueue();
    },

    getCanvas(index) {
      return canvases[index] || null;
    },

    get slideCount() { return n; },
  };
}

function _buildPlaceholder(index, cx, cy, bg) {
  const el = document.createElement('div');
  el.setAttribute('data-placeholder', '1');
  el.style.cssText = `
    position:absolute;inset:0;
    display:flex;align-items:center;justify-content:center;
    background:${bg};
    flex-direction:column;gap:12px;
  `;

  // Animated skeleton lines
  const linesHtml = `
    <div style="width:55%;height:18px;background:#e0e0e0;border-radius:3px;animation:pulse 1.4s ease-in-out infinite;"></div>
    <div style="width:72%;height:10px;background:#ebebeb;border-radius:3px;animation:pulse 1.4s ease-in-out infinite 0.1s;"></div>
    <div style="width:62%;height:10px;background:#ebebeb;border-radius:3px;animation:pulse 1.4s ease-in-out infinite 0.2s;"></div>
    <div style="width:45%;height:10px;background:#ebebeb;border-radius:3px;animation:pulse 1.4s ease-in-out infinite 0.3s;"></div>
  `;

  // Inject keyframes once
  if (!document.getElementById('_pptx_lazy_css')) {
    const style = document.createElement('style');
    style.id = '_pptx_lazy_css';
    style.textContent = `@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}}`;
    document.head.appendChild(style);
  }

  el.innerHTML = linesHtml;
  return el;
}
