/** Voice adapter contract: startListening, stopListening, restartListening, isListening, onCommand, onStatusChange.
 *  Statuses: ready | listening | off | denied | unavailable | unsupported. */
export class WebSpeechAdapter {
  constructor(parse, { cooldownMs, lang }) {
    this.parse = parse; this.cooldownMs = cooldownMs; this.lang = lang;
    this.SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    this.supported = !!this.SR;
    this.status = this.supported ? 'ready' : 'unsupported';
    this.rec = null; this.wanted = false; this.lastAt = 0; this.quickEnds = 0;
    this.cmdCbs = []; this.statusCbs = [];
  }
  onCommand(fn) { this.cmdCbs.push(fn); }
  onStatusChange(fn) { this.statusCbs.push(fn); }
  isListening() { return !!this.rec; }
  _set(s) { if (s !== this.status) { this.status = s; this.statusCbs.forEach(f => f(s)); } }

  startListening() {
    if (!this.supported || this.rec) return;
    this.wanted = true;
    const r = new this.SR(), t0 = Date.now();
    r.lang = this.lang; r.continuous = true; r.interimResults = true;
    r.onstart = () => this._set('listening');
    r.onresult = e => {
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (!e.results[i].isFinal) continue; // ignore interim results
        const cmd = this.parse(e.results[i][0].transcript), now = Date.now();
        if (cmd && now - this.lastAt >= this.cooldownMs) { this.lastAt = now; this.cmdCbs.forEach(f => f(cmd)); }
      }
    };
    r.onerror = e => {
      if (e.error === 'not-allowed' || e.error === 'service-not-allowed') { this.wanted = false; this._set('denied'); }
      else if (e.error === 'audio-capture' || e.error === 'network') { this.wanted = false; this._set('unavailable'); }
      // 'no-speech' / 'aborted': harmless, onend restarts
    };
    r.onend = () => {
      if (this.rec !== r) return; // intentional stop
      this.rec = null;
      this.quickEnds = Date.now() - t0 < 1000 ? this.quickEnds + 1 : 0; // guard against restart loops
      if (this.wanted && this.quickEnds < 5) { this._set('ready'); setTimeout(() => this.wanted && !this.rec && this.startListening(), 300); }
      else if (this.wanted) { this.wanted = false; this._set('unavailable'); }
    };
    this.rec = r;
    try { r.start(); } catch { this.rec = null; this.wanted = false; this._set('unavailable'); }
  }
  stopListening(finalStatus = 'ready') {
    this.wanted = false; this.quickEnds = 0;
    const r = this.rec; this.rec = null;
    if (r) try { r.abort(); } catch { /* already stopped */ }
    if (!this.supported) return;
    if (finalStatus === 'off' || !['denied', 'unavailable'].includes(this.status)) this._set(finalStatus);
  }
  restartListening() { this.stopListening(); this.startListening(); }
}
