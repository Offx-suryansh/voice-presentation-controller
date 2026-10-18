const DEFAULTS = { voiceEnabled: true, defaultZoom: 1, fullscreenOnOpen: false };
export class SettingsManager {
  constructor(storage) { this.storage = storage; }
  get() { return { ...DEFAULTS, ...this.storage.get('settings', {}) }; }
  update(patch) { this.storage.set('settings', { ...this.get(), ...patch }); }
}
