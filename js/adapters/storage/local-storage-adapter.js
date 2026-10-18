export class LocalStorageAdapter {
  constructor(prefix = '') { this.prefix = prefix; }
  get(key, fallback) {
    try { const v = localStorage.getItem(this.prefix + key); return v === null ? fallback : JSON.parse(v); } catch { return fallback; }
  }
  set(key, value) { try { localStorage.setItem(this.prefix + key, JSON.stringify(value)); } catch { /* storage unavailable */ } }
}
