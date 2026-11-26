import { mapLanguageCode } from './languages.js';

const DEFAULT_USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

// In-memory token caches for free Bing/Edge endpoints
let cachedEdgeToken = { value: null, expiresAt: 0 };
let cachedBingWebAuth = {
  host: 'www.bing.com',
  referer: 'https://www.bing.com/translator',
  cookie: '',
  ig: null,
  iid: null,
  key: null,
  token: null,
  count: 0,
  expiresAt: 0
};
let preferredFreeMode = null;

export function resetBingAuthCache() {
  cachedEdgeToken = { value: null, expiresAt: 0 };
  cachedBingWebAuth = {
    host: 'www.bing.com',
    referer: 'https://www.bing.com/translator',
    cookie: '',
    ig: null,
    iid: null,
    key: null,
    token: null,
    count: 0,
    expiresAt: 0
  };
  preferredFreeMode = null;
}

function withTimeoutSignal(timeoutMs = 5000) {
  return typeof AbortSignal !== 'undefined' && typeof AbortSignal.timeout === 'function'
    ? AbortSignal.timeout(timeoutMs)
    : undefined;
}

/**
 * Extract Cookie header string from Fetch Response headers (Node 20 & Cloudflare Workers compatible).
 */
export function extractCookieHeader(headers) {
  if (!headers) return '';
  let rawCookies = [];
  if (typeof headers.getSetCookie === 'function') {
    rawCookies = headers.getSetCookie();
  } else if (typeof headers.get === 'function') {
    const single = headers.get('set-cookie');
    if (single) {
      rawCookies = single.split(/,(?=[^;\s]+=)/);
    }
  }
  return rawCookies
    .map((entry) => String(entry).split(';')[0].trim())
    .filter(Boolean)
    .join('; ');
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
    body: JSON.stringify([{ Text: text }]),
    signal: withTimeoutSignal(10000)
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

  cachedEdgeToken = {
    value: token,
    expiresAt: now + 5 * 60 * 1000
  };
  return token;
}

/**
 * Edge Translate endpoint (fallback when active).
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
 * Fetch Bing web auth parameters + session cookies, following redirects to the effective host.
 */
async function getBingWebAuth(fetchImpl, forceRefresh = false, hostOrder = ['www.bing.com', 'cn.bing.com']) {
  const now = Date.now();
  if (!forceRefresh && cachedBingWebAuth.token && now < cachedBingWebAuth.expiresAt) {
    cachedBingWebAuth.count += 1;
    return cachedBingWebAuth;
  }

  let lastError = null;
  for (const host of hostOrder) {
    try {
      const requestUrl = `https://${host}/translator`;
      const response = await fetchImpl(requestUrl, {
        method: 'GET',
        headers: {
          'User-Agent': DEFAULT_USER_AGENT,
          Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
          'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8'
        },
        signal: withTimeoutSignal(5000)
      });
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      const effectiveUrl = response.url || requestUrl;
      const effectiveHost = new URL(effectiveUrl).host || host;
      const cookie = extractCookieHeader(response.headers);
      const html = await response.text();
      const parsed = parseBingTranslatorHtml(html);

      cachedBingWebAuth = {
        host: effectiveHost,
        referer: effectiveUrl,
        cookie,
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
 * Free Bing engine via Bing Translator web endpoint (/ttranslatev3).
 */
async function translateWithBingWebFree(text, sourceLang, targetLang, fetchImpl) {
  const execute = async (forceRefresh, hostOrder) => {
    const auth = await getBingWebAuth(fetchImpl, forceRefresh, hostOrder);
    const url = `https://${auth.host}/ttranslatev3?isVertical=1&&IG=${encodeURIComponent(auth.ig)}&IID=${encodeURIComponent(auth.iid)}.${auth.count}`;
    const body = new URLSearchParams({
      fromLang: sourceLang,
      text,
      to: targetLang,
      token: auth.token,
      key: auth.key
    });

    const headers = {
      'Content-Type': 'application/x-www-form-urlencoded',
      'User-Agent': DEFAULT_USER_AGENT,
      Accept: 'application/json, text/plain, */*',
      'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
      Origin: `https://${auth.host}`,
      Referer: auth.referer
    };
    if (auth.cookie) {
      headers.Cookie = auth.cookie;
    }

    const response = await fetchImpl(url, {
      method: 'POST',
      headers,
      body: body.toString(),
      signal: withTimeoutSignal(5000)
    });

    const rawText = await response.text();
    let parsed = null;
    try {
      parsed = JSON.parse(rawText);
    } catch {
      parsed = null;
    }
    return { response, rawText, parsed, host: auth.host };
  };

  let attempt = await execute(false, ['www.bing.com', 'cn.bing.com']);

  const hasValidTranslation = (resObj) => {
    if (!resObj.response.ok || !resObj.parsed) return false;
    if (resObj.parsed.statusCode && resObj.parsed.statusCode !== 200) return false;
    const first = Array.isArray(resObj.parsed) ? resObj.parsed[0] : resObj.parsed;
    return typeof first?.translations?.[0]?.text === 'string';
  };

  if (!hasValidTranslation(attempt)) {
    // Refresh session and try alternate Bing host first
    const altOrder = attempt.host.includes('cn.bing.com')
      ? ['www.bing.com', 'cn.bing.com']
      : ['cn.bing.com', 'www.bing.com'];
    attempt = await execute(true, altOrder);
  }

  if (!hasValidTranslation(attempt)) {
    const snippet = (attempt.rawText || '').replace(/\s+/g, ' ').slice(0, 180);
    throw new Error(
      `Bing Web Translate failed (${attempt.host}, HTTP ${attempt.response.status}): ${snippet || 'empty response'}`
    );
  }

  const first = Array.isArray(attempt.parsed) ? attempt.parsed[0] : attempt.parsed;
  return {
    translatedText: first.translations[0].text,
    detectedSourceLanguage: first?.detectedLanguage?.language || sourceLang,
    mode: 'free-web'
  };
}

/**
 * Translate text with Bing Translator (official API key or free Web/Edge fallback).
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

  // Prefer Bing Web endpoint first since Edge anonymous auth returns 404 on many edges
  const order = preferredFreeMode === 'edge' ? ['edge', 'web'] : ['web', 'edge'];
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
