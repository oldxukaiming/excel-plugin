import { mapLanguageCode } from './languages.js';

const DEFAULT_USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

function withTimeoutSignal(timeoutMs = 8000) {
  return typeof AbortSignal !== 'undefined' && typeof AbortSignal.timeout === 'function'
    ? AbortSignal.timeout(timeoutMs)
    : undefined;
}

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
    body: JSON.stringify(payload),
    signal: withTimeoutSignal(10000)
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
 * Parse translate_a/single array format: [[[translated, source], ...], null, detectedLang]
 */
function parseSingleApiPayload(data, sourceLang) {
  if (!Array.isArray(data) || !Array.isArray(data[0])) {
    return null;
  }
  const translatedText = data[0]
    .filter((segment) => Array.isArray(segment) && typeof segment[0] === 'string')
    .map((segment) => segment[0])
    .join('');

  return {
    translatedText,
    detectedSourceLanguage: typeof data[2] === 'string' ? data[2] : sourceLang
  };
}

/**
 * Parse clients5 /translate_a/t payload:
 * - sl=auto: [["translated text", "en"]]
 * - sl=en:   ["translated text"]
 */
function parseClients5Payload(data, sourceLang) {
  if (!Array.isArray(data) || data.length === 0) {
    return null;
  }
  const first = data[0];
  if (typeof first === 'string') {
    return {
      translatedText: data.filter((item) => typeof item === 'string').join(''),
      detectedSourceLanguage: sourceLang
    };
  }
  if (Array.isArray(first) && typeof first[0] === 'string') {
    return {
      translatedText: data
        .filter((item) => Array.isArray(item) && typeof item[0] === 'string')
        .map((item) => item[0])
        .join(''),
      detectedSourceLanguage: typeof first[1] === 'string' ? first[1] : sourceLang
    };
  }
  return null;
}

async function fetchSingleEndpoint(baseUrl, clientName, text, sourceLang, targetLang, fetchImpl) {
  const params = new URLSearchParams({
    client: clientName,
    sl: sourceLang,
    tl: targetLang,
    dt: 't',
    q: text
  });
  const url = `${baseUrl}/translate_a/single?${params.toString()}`;
  const response = await fetchImpl(url, {
    method: 'GET',
    headers: { 'User-Agent': DEFAULT_USER_AGENT },
    signal: withTimeoutSignal(8000)
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`${clientName} HTTP ${response.status}: ${errText.slice(0, 120)}`);
  }

  const data = await response.json();
  const parsed = parseSingleApiPayload(data, sourceLang);
  if (!parsed) {
    throw new Error(`Unexpected response format from ${baseUrl} (${clientName})`);
  }
  return { ...parsed, mode: 'free-web' };
}

async function fetchClients5Endpoint(text, sourceLang, targetLang, fetchImpl) {
  const params = new URLSearchParams({
    client: 'dict-chrome-ex',
    sl: sourceLang,
    tl: targetLang,
    q: text
  });
  const url = `https://clients5.google.com/translate_a/t?${params.toString()}`;
  const response = await fetchImpl(url, {
    method: 'GET',
    headers: { 'User-Agent': DEFAULT_USER_AGENT },
    signal: withTimeoutSignal(8000)
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`clients5 HTTP ${response.status}: ${errText.slice(0, 120)}`);
  }

  const data = await response.json();
  const parsed = parseClients5Payload(data, sourceLang);
  if (!parsed) {
    throw new Error('Unexpected response format from clients5.google.com');
  }
  return { ...parsed, mode: 'free-web' };
}

/**
 * Translate text using Google Translate free public endpoints (rotates endpoints on 429/failure).
 */
async function translateWithFreeEndpoints(text, sourceLang, targetLang, customBaseUrl, fetchImpl) {
  if (customBaseUrl && customBaseUrl.trim()) {
    const cleaned = customBaseUrl.trim().replace(/\/+$/, '');
    return fetchSingleEndpoint(cleaned, 'gtx', text, sourceLang, targetLang, fetchImpl);
  }

  const strategies = [
    () => fetchClients5Endpoint(text, sourceLang, targetLang, fetchImpl),
    () => fetchSingleEndpoint('https://translate.googleapis.com', 'dict-chrome-ex', text, sourceLang, targetLang, fetchImpl),
    () => fetchSingleEndpoint('https://translate.googleapis.com', 'gtx', text, sourceLang, targetLang, fetchImpl)
  ];

  const errors = [];
  for (const run of strategies) {
    try {
      return await run();
    } catch (err) {
      errors.push(err.message || 'Unknown error');
    }
  }

  throw new Error(`Google Free Translate error: ${errors.join(' | ')}`);
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

  const result = await translateWithFreeEndpoints(text, sourceLang, targetLang, customEndpoint, fetchImpl);
  return {
    engine: 'google',
    sourceLanguage: sourceLang,
    targetLanguage: targetLang,
    ...result
  };
}
