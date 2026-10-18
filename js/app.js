import { APP_CONFIG as C } from './config.js';
import { createState } from './core/presentation-state.js';
import { PresentationController } from './core/presentation-controller.js';
import { parseCommand } from './core/command-parser.js';
import { HistoryManager } from './core/history-manager.js';
import { SettingsManager } from './core/settings-manager.js';
import { LocalStorageAdapter } from './adapters/storage/local-storage-adapter.js';
import { BrowserPlatform } from './adapters/platform/browser-platform.js';
import { WebSpeechAdapter } from './adapters/voice/web-speech-adapter.js';
import { PresentationViewer } from './viewer/presentation-viewer.js';

const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
$('#copyrightYear').textContent = new Date().getFullYear();
const storage = new LocalStorageAdapter(C.storageKey);
const settings = new SettingsManager(storage), history = new HistoryManager(storage, C.history.max);
const platform = new BrowserPlatform($('#viewer')), voice = new WebSpeechAdapter(parseCommand, C.voice);
const viewer = new PresentationViewer($('#canvas'), $('#stage'));
const initial = settings.get();
const state = createState({
  currentSlide: 1, totalSlides: 0, zoom: initial.defaultZoom, fileName: '', isFullscreen: false,
  voiceEnabled: initial.voiceEnabled, voiceStatus: !voice.supported ? 'unsupported' : initial.voiceEnabled ? 'ready' : 'off'
});
const ctl = new PresentationController(state, platform, () => settings.get().defaultZoom, C.presentation);

/* ---- voice status ---- */
const MIC = 'Microphone access is unavailable. You can still control slides using keyboard controls.';
const VOICE_TEXT = { ready: '● Voice Ready', listening: '● Listening...', off: '○ Voice Off', denied: MIC, unavailable: MIC,
  unsupported: "Voice control isn't supported in this browser. Keyboard controls are still available." };
let flashMsg = '', flashTimer = 0, lastKey = '';
function renderVoice() {
  const st = state.get().voiceStatus, text = flashMsg || VOICE_TEXT[st];
  const cls = flashMsg || st === 'listening' ? 'ok' : ['denied', 'unavailable', 'unsupported'].includes(st) ? 'warn' : '';
  $$('.vstatus').forEach(el => { el.textContent = text; el.className = 'vstatus ' + cls; });
}
function flash(msg) {
  flashMsg = msg; renderVoice(); clearTimeout(flashTimer);
  flashTimer = setTimeout(() => { flashMsg = ''; renderVoice(); }, 1500);
}
function toast(msg) {
  const t = Object.assign(document.createElement('div'), { className: 'toast', textContent: msg });
  t.setAttribute('role', 'alert'); document.body.append(t); setTimeout(() => t.remove(), 4000);
}

/* ---- render from state ---- */
state.subscribe(s => {
  const pct = ctl.progress();
  $('#slideText').textContent = `Slide ${s.currentSlide} / ${s.totalSlides}`;
  $('#fill').style.width = pct + '%'; $('#bar').setAttribute('aria-valuenow', Math.round(pct));
  $('#zoomText').textContent = Math.round(s.zoom * 100) + '%'; $('#fname').textContent = s.fileName;
  renderVoice();
  const key = [s.currentSlide, s.zoom, s.isFullscreen, s.totalSlides].join();
  if (key !== lastKey && viewer.ready && !$('#viewer').hidden) {
    lastKey = key;
    viewer.render(s.currentSlide, s.zoom).catch(() => toast('This slide could not be rendered. Try reopening the presentation.'));
    history.record(s.fileName, s.currentSlide);
  }
});

/* ---- open / close ---- */
async function openFile(file) {
  try {
    const pages = await viewer.load(file), last = history.get(file.name);
    $('#viewer').hidden = false; lastKey = '';
    state.set({ fileName: file.name, totalSlides: pages, zoom: settings.get().defaultZoom,
      currentSlide: last && last.lastSlide <= pages ? last.lastSlide : 1 });
    if (settings.get().fullscreenOnOpen && !document.fullscreenElement) platform.toggleFullscreen();
    if (state.get().voiceEnabled) voice.startListening();
  } catch (e) {
    toast(e.code === 'invalid'
      ? 'Choose a valid PDF or PowerPoint (.pptx) file.'
      : "We couldn't open this presentation. Check the file and try again.");
  }
}
function closeViewer() {
  voice.stopListening(state.get().voiceEnabled ? 'ready' : 'off'); platform.exitFullscreen();
  viewer.close(); $('#viewer').hidden = true; state.set({ totalSlides: 0, fileName: '' }); renderHistory();
}

/* ---- history & settings views ---- */
function renderHistory() {
  const box = $('#historyList'), items = history.list(); box.textContent = '';
  if (!items.length) return void (box.innerHTML = '<span class="muted">No presentations yet.</span>');
  for (const h of items) {
    const row = Object.assign(document.createElement('div'), { className: 'hist' });
    const a = Object.assign(document.createElement('span'), { textContent: h.fileName });
    const b = Object.assign(document.createElement('span'), { className: 'muted small',
      textContent: `Slide ${h.lastSlide} · ${new Date(h.lastOpened).toLocaleString()} · select the file again to reopen` });
    row.append(a, b); box.append(row);
  }
}
function showView(name) {
  $$('.view').forEach(v => v.classList.toggle('active', v.id === 'view-' + name));
  $$('nav [data-view]').forEach(b => b.classList.toggle('active', b.dataset.view === name));
}

/* ---- wiring ---- */
const pick = () => $('#file').click();
$('#openBtn').onclick = $('#navOpen').onclick = pick;
$('#file').onchange = e => { openFile(e.target.files[0]); e.target.value = ''; };
$$('nav [data-view]').forEach(b => b.onclick = () => showView(b.dataset.view));
$('#backBtn').onclick = closeViewer;
$('#prevBtn').onclick = () => ctl.previousSlide(); $('#nextBtn').onclick = () => ctl.nextSlide();
$('#zoomIn').onclick = () => ctl.zoomIn(); $('#zoomOut').onclick = () => ctl.zoomOut();
$('#fsBtn').onclick = $('#fsTop').onclick = () => ctl.toggleFullscreen();
$('#bar').onclick = e => {
  const r = e.currentTarget.getBoundingClientRect();
  ctl.goToSlide(Math.ceil((e.clientX - r.left) / r.width * state.get().totalSlides));
};
$('#clearHistory').onclick = () => { history.clear(); renderHistory(); };
platform.onFullscreenChange(isFullscreen => state.set({ isFullscreen }));
window.addEventListener('resize', () => { lastKey = ''; state.set({}); });

const drop = $('#drop');
['dragenter', 'dragover'].forEach(t => window.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(t => window.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
window.addEventListener('drop', e => e.dataTransfer.files[0] && openFile(e.dataTransfer.files[0]));

const keys = { arrowright: () => ctl.nextSlide(), arrowleft: () => ctl.previousSlide(), '+': () => ctl.zoomIn(), '=': () => ctl.zoomIn(),
  '-': () => ctl.zoomOut(), '0': () => ctl.resetZoom(), f: () => ctl.toggleFullscreen(), escape: () => platform.exitFullscreen() };
document.addEventListener('keydown', e => {
  if ($('#viewer').hidden || /INPUT|SELECT|TEXTAREA/.test(e.target.tagName) || e.ctrlKey || e.metaKey || e.altKey) return;
  const act = keys[e.key.toLowerCase()];
  if (act) { e.preventDefault(); act(); }
});

voice.onStatusChange(voiceStatus => state.set({ voiceStatus }));
voice.onCommand(cmd => {
  if (cmd === 'next') { ctl.nextSlide(); flash('✓ Next Slide'); } else { ctl.previousSlide(); flash('✓ Previous Slide'); }
});

$('#setVoice').checked = initial.voiceEnabled; $('#setZoom').value = Math.round(initial.defaultZoom * 100); $('#setFs').checked = initial.fullscreenOnOpen;
$('#setVoice').onchange = e => {
  const on = e.target.checked; settings.update({ voiceEnabled: on });
  state.set({ voiceEnabled: on, voiceStatus: !voice.supported ? 'unsupported' : on ? 'ready' : 'off' });
  if (!$('#viewer').hidden) on ? voice.startListening() : voice.stopListening('off');
};
$('#setZoom').onchange = e => settings.update({ defaultZoom: +e.target.value / 100 });
$('#setFs').onchange = e => settings.update({ fullscreenOnOpen: e.target.checked });
platform.listMicrophones().then(m => { if (m.length > 1) $('#setMic').innerHTML = m.map((d, i) => `<option>${d.label || 'Microphone ' + (i + 1)}</option>`).join(''); });
renderHistory(); renderVoice();
