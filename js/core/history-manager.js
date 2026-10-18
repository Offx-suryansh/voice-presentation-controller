/** Metadata only: fileName, lastOpened, lastSlide. Never file contents. */
export class HistoryManager {
  constructor(storage, max) { this.storage = storage; this.max = max; }
  list() { return this.storage.get('history', []); }
  get(fileName) { return this.list().find(h => h.fileName === fileName); }
  record(fileName, lastSlide) {
    const rest = this.list().filter(h => h.fileName !== fileName);
    this.storage.set('history', [{ fileName, lastOpened: Date.now(), lastSlide }, ...rest].slice(0, this.max));
  }
  clear() { this.storage.set('history', []); }
}
