import { mapLanguageCode } from './languages.js';

const DEFAULT_USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

/**
 * Translate text using Google Cloud Translation v2 official API.
 */
async function translateWithOfficialApi(text, sourceLang, targetLang, apiKey, fetchImpl) {
  const url = `https://translation.googleapis.com/language/translate/v2?key=${encodeURIComponent(apiKey)}`;
  const payload = {
    q: text,
    target: targetLang,
    format: 'text'
  };
  if (sourceLang && sourceLang !== 'auto') {
    payload.source = sourceLang;
  }

  const response = await fetchImpl(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Google Official API error (${response.status}): ${errText}`);
  }

  const data = await response.json();
  const translation = data?.data?.translations?.[0];
  if (!translation || typeof translation.translatedText !== 'string') {
    throw new Error('Unexpected response format from Google Official API');
  }

  return {
    translatedText: translation.translatedText,
    detectedSourceLanguage: translation.detectedSourceLanguage || sourceLang,
    mode: 'official-api'
  };
}

/**
 * Translate text using Google Translate public web endpoint (no API key required).
 */
async function translateWithFreeEndpoint(text, sourceLang, targetLang, customBaseUrl, fetchImpl) {
  const baseUrl = (customBaseUrl || 'https://translate.googleapis.com').replace(/\/+$/, '');
  const params = new URLSearchParams({
    client: 'gtx',
    sl: sourceLang,
    tl: targetLang,
    dt: 't',
    q: text
  });

  const url = `${baseUrl}/translate_a/single?${params.toString()}`;
  const signal =
    typeof AbortSignal !== 'undefined' && typeof AbortSignal.timeout === 'function'
      ? AbortSignal.timeout(8000)
      : undefined;
  const response = await fetchImpl(url, {
    method: 'GET',
    headers: {
      'User-Agent': DEFAULT_USER_AGENT
    },
    signal
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Google Free Translate error (${response.status}): ${errText.slice(0, 200)}`);
  }

  const data = await response.json();
  if (!Array.isArray(data) || !Array.isArray(data[0])) {
    throw new Error('Unexpected response format from Google Free Translate');
  }

  const translatedText = data[0]
    .filter((segment) => Array.isArray(segment) && typeof segment[0] === 'string')
    .map((segment) => segment[0])
    .join('');

  return {
    translatedText,
    detectedSourceLanguage: typeof data[2] === 'string' ? data[2] : sourceLang,
    mode: 'free-web'
  };
}

/**
 * Translate text with Google Translate (supports both free web mode and official API key).
 * @param {object} options
 * @param {string} options.text
 * @param {string} [options.from='auto']
 * @param {string} [options.to='zh-CN']
 * @param {string} [options.apiKey]
 * @param {string} [options.customEndpoint]
 * @param {typeof fetch} [options.fetchImpl=fetch]
 */
export async function translateWithGoogle({
  text,
  from = 'auto',
  to = 'zh-CN',
  apiKey,
  customEndpoint,
  fetchImpl = fetch
}) {
  const sourceLang = mapLanguageCode(from, 'google', true);
  const targetLang = mapLanguageCode(to, 'google', false);

  if (apiKey && apiKey.trim()) {
    const result = await translateWithOfficialApi(text, sourceLang, targetLang, apiKey.trim(), fetchImpl);
    return {
      engine: 'google',
      sourceLanguage: sourceLang,
      targetLanguage: targetLang,
      ...result
    };
  }

  const result = await translateWithFreeEndpoint(text, sourceLang, targetLang, customEndpoint, fetchImpl);
  return {
    engine: 'google',
    sourceLanguage: sourceLang,
    targetLanguage: targetLang,
    ...result
  };
}
