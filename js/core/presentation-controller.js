/** All navigation/zoom goes through here: voice, keyboard, buttons, progress bar. */
export class PresentationController {
  constructor(state, platform, getDefaultZoom, { minZoom, maxZoom, zoomStep }) {
    Object.assign(this, { state, platform, getDefaultZoom, minZoom, maxZoom, zoomStep });
  }
  goToSlide(i) {
    const { currentSlide, totalSlides } = this.state.get();
    const n = Math.min(Math.max(Math.round(i), 1), totalSlides);
    if (totalSlides && n !== currentSlide) this.state.set({ currentSlide: n });
  }
  nextSlide() { this.goToSlide(this.state.get().currentSlide + 1); }
  previousSlide() { this.goToSlide(this.state.get().currentSlide - 1); }
  setZoom(z) { this.state.set({ zoom: Math.round(Math.min(Math.max(z, this.minZoom), this.maxZoom) * 100) / 100 }); }
  zoomIn() { this.setZoom(this.state.get().zoom + this.zoomStep); }
  zoomOut() { this.setZoom(this.state.get().zoom - this.zoomStep); }
  resetZoom() { this.setZoom(this.getDefaultZoom()); }
  toggleFullscreen() { this.platform.toggleFullscreen(); }
  /** currentSlide / totalSlides × 100, clamped to 0–100. */
  progress() {
    const { currentSlide, totalSlides } = this.state.get();
    return totalSlides ? Math.min(100, Math.max(0, currentSlide / totalSlides * 100)) : 0;
  }
}
