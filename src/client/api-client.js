/**
 * Shared client API helper supporting same-origin requests,
 * Cloudflare Worker URLs configured via config.js, or user-saved apiBaseUrl.
 */

export const STORAGE_KEY = 'excel-translator-settings-v1';

export function getDefaultSettings() {
  return {
    apiBaseUrl: '',
    bingApiKey: '',
    bingRegion: '',
    googleApiKey: '',
    googleEndpoint: ''
  };
}

export function loadTranslatorState() {
  const defaults = {
    engine: 'bing',
    from: 'auto',
    to: 'zh-CN',
    targetMode: 'below',
    settings: getDefaultSettings()
  };

  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaults;
    const parsed = JSON.parse(raw);
    return {
      engine: parsed.engine === 'google' ? 'google' : 'bing',
      from: parsed.from || 'auto',
      to: parsed.to || 'zh-CN',
      targetMode: parsed.targetMode || 'below',
      settings: { ...defaults.settings, ...(parsed.settings || {}) }
    };
  } catch {
    return defaults;
  }
}

export function saveTranslatorState(state) {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        engine: state.engine,
        from: state.from,
        to: state.to,
        targetMode: state.targetMode || 'below',
        settings: state.settings
      })
    );
  } catch {
    // Ignore storage quota errors
  }
}

export function resolveApiBaseUrl(settings = {}) {
  const userBase = (settings.apiBaseUrl || '').trim().replace(/\/+$/, '');
  if (userBase) return userBase;

  const runtimeBase = (globalThis.__EXCEL_TRANSLATOR_CONFIG__?.apiBaseUrl || '')
    .trim()
    .replace(/\/+$/, '');
  if (runtimeBase) return runtimeBase;

  return '';
}

export function buildApiUrl(path, settings = {}) {
  const base = resolveApiBaseUrl(settings);
  return `${base}${path}`;
}

export async function apiFetch(path, options = {}, settings = {}) {
  const url = buildApiUrl(path, settings);
  return fetch(url, options);
}
