import { APP_CONFIG } from '../config.js';
export const normalize = t => t.toLowerCase().replace(/[.,!?]/g, '').trim().replace(/\s+/g, ' ');
/** Returns 'next' | 'previous' | null. Exact phrase matches only. */
export function parseCommand(transcript) {
  const n = normalize(transcript), v = APP_CONFIG.voice;
  return n === v.nextCommand ? 'next' : n === v.previousCommand ? 'previous' : null;
}
