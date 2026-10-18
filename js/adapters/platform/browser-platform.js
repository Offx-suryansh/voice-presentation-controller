/** Browser-specific platform features; swap for Electron/Capacitor equivalents later. */
export class BrowserPlatform {
  constructor(fullscreenElement) { this.el = fullscreenElement; }
  toggleFullscreen() {
    if (document.fullscreenElement) return document.exitFullscreen();
    return this.el.requestFullscreen?.().catch(() => {});
  }
  exitFullscreen() { if (document.fullscreenElement) document.exitFullscreen(); }
  onFullscreenChange(fn) { document.addEventListener('fullscreenchange', () => fn(!!document.fullscreenElement)); }
  async listMicrophones() {
    try { return (await navigator.mediaDevices.enumerateDevices()).filter(d => d.kind === 'audioinput'); } catch { return []; }
  }
}
