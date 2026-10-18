// Central configuration: change behaviour here, not in feature code.
export const APP_CONFIG = {
  voice: { nextCommand: 'next slide', previousCommand: 'previous slide', cooldownMs: 1000, lang: 'en-US' },
  presentation: { defaultZoom: 1, minZoom: 0.25, maxZoom: 3, zoomStep: 0.1 },
  history: { max: 20 },
  storageKey: 'vpc.'
};
