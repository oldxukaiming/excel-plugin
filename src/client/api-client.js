/**
 * Shared client API helper supporting same-origin requests,
 * Cloudflare Worker URLs configured via config.js, user-saved apiBaseUrl,
 * and Bearer session token authentication.
 */

export const STORAGE_KEY = 'excel-translator-settings-v1';
export const AUTH_STORAGE_KEY = 'excel-translator-auth-v1';

let unauthorizedHandler = null;

export function setUnauthorizedHandler(handler) {
  unauthorizedHandler = typeof handler === 'function' ? handler : null;
}

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

export function loadAuthSession() {
  try {
    const raw = localStorage.getItem(AUTH_STORAGE_KEY);
    if (!raw) return { token: '', username: '' };
    const parsed = JSON.parse(raw);
    return {
      token: String(parsed.token || ''),
      username: String(parsed.username || '')
    };
  } catch {
    return { token: '', username: '' };
  }
}

export function saveAuthSession({ token = '', username = '' }) {
  try {
    if (!token) {
      localStorage.removeItem(AUTH_STORAGE_KEY);
      return;
    }
    localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify({ token, username }));
  } catch {
    // Ignore storage quota errors
  }
}

export function clearAuthSession() {
  saveAuthSession({ token: '', username: '' });
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
  const { token } = loadAuthSession();
  const headers = new Headers(options.headers || {});
  if (token && !headers.has('Authorization')) {
    headers.set('Authorization', `Bearer ${token}`);
  }

  const response = await fetch(url, { ...options, headers });
  if (response.status === 401 && path !== '/api/auth/login') {
    clearAuthSession();
    if (unauthorizedHandler) {
      unauthorizedHandler();
    }
  }
  return response;
}

export async function fetchAuthStatus(settings = {}) {
  const res = await apiFetch('/api/auth/status', {}, settings);
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  return res.json();
}

export async function loginWithCredentials(username, password, settings = {}) {
  const res = await apiFetch(
    '/api/auth/login',
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    },
    settings
  );
  const data = await res.json();
  if (!res.ok || !data.ok) {
    throw new Error(data.error || '用户名或密码错误');
  }
  if (data.token) {
    saveAuthSession({ token: data.token, username: data.username || username });
  }
  return data;
}
