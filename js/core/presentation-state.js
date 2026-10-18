/** Single source of truth; UI subscribes instead of keeping its own copies. */
export function createState(initial) {
  let s = { ...initial }; const subs = new Set();
  return {
    get: () => s,
    set(patch) { s = { ...s, ...patch }; subs.forEach(f => f(s)); },
    subscribe(f) { subs.add(f); return () => subs.delete(f); }
  };
}
