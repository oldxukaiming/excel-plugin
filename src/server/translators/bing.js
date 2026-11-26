import { mapLanguageCode } from './languages.js';

const DEFAULT_USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

// In-memory token caches for free Bing/Edge endpoints
let cachedEdgeToken = { value: null, expiresAt: 0 };
let cachedBingWebAuth = { host: 'www.bing.com', ig: null, iid: null, key: null, token: null, count: 0, expiresAt: 0 };
let preferredFreeMode = null;

export function resetBingAuthCache() {
  cachedEdgeToken = { value: null, expiresAt: 0 };
  cachedBingWebAuth = { host: 'www.bing.com', ig: null, iid: null, key: null, token: null, count: 0, expiresAt: 0 };
  preferredFreeMode = null;
}

function withTimeoutSignal(timeoutMs = 5000) {
  return typeof AbortSignal !== 'undefined' && typeof AbortSignal.timeout === 'function'
    ? AbortSignal.timeout(timeoutMs)
    : undefined;
}

/**
 * Translate text using Azure Cognitive Services Translator official API.
 */
async function translateWithOfficialApi(text, sourceLang, targetLang, apiKey, region, fetchImpl) {
  const params = new URLSearchParams({
    'api-version': '3.0',
    to: targetLang
  });
  if (sourceLang && sourceLang !== 'auto-detect') {
    params.set('from', sourceLang);
  }

  const headers = {
    'Ocp-Apim-Subscription-Key': apiKey,
    'Content-Type': 'application/json'
  };
  if (region && region.trim()) {
    headers['Ocp-Apim-Subscription-Region'] = region.trim();
  }

  const url = `https://api.cognitive.microsofttranslator.com/translate?${params.toString()}`;
  const response = await fetchImpl(url, {
    method: 'POST',
    headers,
    body: JSON.stringify([{ Text: text }])
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Bing/Azure Official API error (${response.status}): ${errText}`);
  }

  const data = await response.json();
  const first = data?.[0];
  const translated = first?.translations?.[0]?.text;
  if (typeof translated !== 'string') {
    throw new Error('Unexpected response format from Bing/Azure Official API');
  }

  return {
    translatedText: translated,
    detectedSourceLanguage: first?.detectedLanguage?.language || sourceLang,
    mode: 'official-api'
  };
}

/**
 * Fetch short-lived Edge Translate JWT token (free, no key required).
 */
async function getEdgeAuthToken(fetchImpl, forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && cachedEdgeToken.value && now < cachedEdgeToken.expiresAt) {
    return cachedEdgeToken.value;
  }

  const response = await fetchImpl('https://edge.microsoft.com/translate/auth', {
    method: 'GET',
    headers: { 'User-Agent': DEFAULT_USER_AGENT },
    signal: withTimeoutSignal(4000)
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch Edge translate auth token (${response.status})`);
  }

  const token = (await response.text()).trim();
  if (!token) {
    throw new Error('Empty token received from Edge translate auth');
  }

  // Edge JWT typically lasts ~10 minutes; refresh after 5 minutes to stay safe
  cachedEdgeToken = {
    value: token,
    expiresAt: now + 5 * 60 * 1000
  };
  return token;
}

/**
 * Primary free Bing engine via Microsoft Edge Translate endpoint.
 */
async function translateWithEdgeFree(text, sourceLang, targetLang, fetchImpl) {
  const execute = async (forceRefresh) => {
    const jwt = await getEdgeAuthToken(fetchImpl, forceRefresh);
    const params = new URLSearchParams({
      to: targetLang,
      'api-version': '3.0',
      includeSentenceLength: 'true'
    });
    if (sourceLang && sourceLang !== 'auto-detect') {
      params.set('from', sourceLang);
    }

    const url = `https://api-edge.cognitive.microsofttranslator.com/translate?${params.toString()}`;
    return fetchImpl(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${jwt}`,
        'Content-Type': 'application/json',
        'User-Agent': DEFAULT_USER_AGENT
      },
      body: JSON.stringify([{ Text: text }]),
      signal: withTimeoutSignal(5000)
    });
  };

  let response = await execute(false);
  if (response.status === 401 || response.status === 403) {
    response = await execute(true);
  }

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Bing Edge free endpoint error (${response.status}): ${errText.slice(0, 200)}`);
  }

  const data = await response.json();
  const first = data?.[0];
  const translated = first?.translations?.[0]?.text;
  if (typeof translated !== 'string') {
    throw new Error('Unexpected response format from Bing Edge endpoint');
  }

  return {
    translatedText: translated,
    detectedSourceLanguage: first?.detectedLanguage?.language || sourceLang,
    mode: 'free-edge'
  };
}

/**
 * Parse IG, IID, and abuse prevention helper token from Bing Translator HTML.
 */
export function parseBingTranslatorHtml(html) {
  const igMatch = html.match(/IG:"([A-Za-z0-9]+)"/);
  const iidMatch = html.match(/data-iid="([^"]+)"/) || html.match(/IID=([A-Za-z0-9._]+)/);
  const paramsMatch = html.match(/var params_AbusePreventionHelper\s*=\s*\[(\d+),"([^"]+)",(\d+)\]/);

  if (!igMatch || !paramsMatch) {
    throw new Error('Unable to parse Bing Translator tokens from page HTML');
  }

  return {
    ig: igMatch[1],
    iid: iidMatch ? iidMatch[1] : 'translator.5024',
    key: paramsMatch[1],
    token: paramsMatch[2],
    ttlMs: Number(paramsMatch[3]) || 1800000
  };
}

/**
 * Fetch Bing web auth parameters from cn.bing.com or www.bing.com.
 */
async function getBingWebAuth(fetchImpl, forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && cachedBingWebAuth.token && now < cachedBingWebAuth.expiresAt) {
    cachedBingWebAuth.count += 1;
    return cachedBingWebAuth;
  }

  const hosts = ['cn.bing.com', 'www.bing.com'];
  let lastError = null;

  for (const host of hosts) {
    try {
      const response = await fetchImpl(`https://${host}/translator`, {
        method: 'GET',
        headers: { 'User-Agent': DEFAULT_USER_AGENT },
        signal: withTimeoutSignal(5000)
      });
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      const html = await response.text();
      const parsed = parseBingTranslatorHtml(html);
      cachedBingWebAuth = {
        host,
        ig: parsed.ig,
        iid: parsed.iid,
        key: parsed.key,
        token: parsed.token,
        count: 1,
        expiresAt: now + Math.min(parsed.ttlMs, 15 * 60 * 1000)
      };
      return cachedBingWebAuth;
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(`Failed to initialize Bing web translator: ${lastError?.message || 'unknown error'}`);
}

/**
 * Fallback free Bing engine via Bing Translator web endpoint (/ttranslatev3).
 */
async function translateWithBingWebFree(text, sourceLang, targetLang, fetchImpl) {
  const execute = async (forceRefresh) => {
    const auth = await getBingWebAuth(fetchImpl, forceRefresh);
    const url = `https://${auth.host}/ttranslatev3?isVertical=1&&IG=${encodeURIComponent(auth.ig)}&IID=${encodeURIComponent(auth.iid)}.${auth.count}`;
    const body = new URLSearchParams({
      fromLang: sourceLang,
      text,
      to: targetLang,
      token: auth.token,
      key: auth.key
    });

    return fetchImpl(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'User-Agent': DEFAULT_USER_AGENT,
        Referer: `https://${auth.host}/translator`
      },
      body: body.toString(),
      signal: withTimeoutSignal(5000)
    });
  };

  let response = await execute(false);
  let data = response.ok ? await response.json().catch(() => null) : null;

  if (!response.ok || (data && data.statusCode && data.statusCode !== 200)) {
    response = await execute(true);
    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Bing Web Translate error (${response.status}): ${errText.slice(0, 200)}`);
    }
    data = await response.json();
  }

  const first = Array.isArray(data) ? data[0] : data;
  const translated = first?.translations?.[0]?.text;
  if (typeof translated !== 'string') {
    throw new Error('Unexpected response format from Bing Web Translate');
  }

  return {
    translatedText: translated,
    detectedSourceLanguage: first?.detectedLanguage?.language || sourceLang,
    mode: 'free-web'
  };
}

/**
 * Translate text with Bing Translator (official API key or free Edge/Web fallback).
 * @param {object} options
 * @param {string} options.text
 * @param {string} [options.from='auto']
 * @param {string} [options.to='zh-CN']
 * @param {string} [options.apiKey]
 * @param {string} [options.region]
 * @param {typeof fetch} [options.fetchImpl=fetch]
 */
export async function translateWithBing({
  text,
  from = 'auto',
  to = 'zh-CN',
  apiKey,
  region,
  fetchImpl = fetch
}) {
  const sourceLang = mapLanguageCode(from, 'bing', true);
  const targetLang = mapLanguageCode(to, 'bing', false);

  if (apiKey && apiKey.trim()) {
    const result = await translateWithOfficialApi(text, sourceLang, targetLang, apiKey.trim(), region, fetchImpl);
    return {
      engine: 'bing',
      sourceLanguage: sourceLang,
      targetLanguage: targetLang,
      ...result
    };
  }

  const order = preferredFreeMode === 'web' ? ['web', 'edge'] : ['edge', 'web'];
  const errors = [];

  for (const mode of order) {
    try {
      const result =
        mode === 'edge'
          ? await translateWithEdgeFree(text, sourceLang, targetLang, fetchImpl)
          : await translateWithBingWebFree(text, sourceLang, targetLang, fetchImpl);
      preferredFreeMode = mode;
      return {
        engine: 'bing',
        sourceLanguage: sourceLang,
        targetLanguage: targetLang,
        ...result
      };
    } catch (err) {
      errors.push(`${mode}: [${err?.message}]`);
    }
  }

  throw new Error(`Bing free translation failed. ${errors.join('; ')}`);
}
