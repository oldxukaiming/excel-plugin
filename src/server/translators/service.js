import { translateWithBing } from './bing.js';
import { translateWithGoogle } from './google.js';

const MAX_CACHE_SIZE = 1000;
const translationCache = new Map();

export function clearTranslationCache() {
  translationCache.clear();
}

function buildCacheKey({ engine, from, to, text, hasApiKey }) {
  return `${engine}|${from}|${to}|${hasApiKey ? 'key' : 'free'}|${text}`;
}

function setCache(key, value) {
  if (translationCache.size >= MAX_CACHE_SIZE) {
    const oldestKey = translationCache.keys().next().value;
    translationCache.delete(oldestKey);
  }
  translationCache.set(key, value);
}

/**
 * Translate a single string using the selected engine ('bing' or 'google').
 */
export async function translateText({
  text,
  engine = 'bing',
  from = 'auto',
  to = 'zh-CN',
  config = {},
  fetchImpl = fetch
}) {
  if (text === null || text === undefined) {
    return {
      engine,
      translatedText: '',
      detectedSourceLanguage: from,
      mode: 'passthrough'
    };
  }

  const rawText = String(text);
  if (!rawText.trim()) {
    return {
      engine,
      translatedText: rawText,
      detectedSourceLanguage: from,
      mode: 'passthrough'
    };
  }

  const normalizedEngine = String(engine).toLowerCase();
  if (normalizedEngine !== 'bing' && normalizedEngine !== 'google') {
    throw new Error(`Unsupported translation engine: ${engine}. Expected 'bing' or 'google'.`);
  }

  const runtimeEnv = typeof process !== 'undefined' && process?.env ? process.env : {};
  const googleApiKey = config.googleApiKey || runtimeEnv.GOOGLE_TRANSLATE_API_KEY || '';
  const googleEndpoint = config.googleEndpoint || runtimeEnv.GOOGLE_TRANSLATE_ENDPOINT || '';
  const bingApiKey = config.bingApiKey || runtimeEnv.BING_TRANSLATOR_API_KEY || '';
  const bingRegion = config.bingRegion || runtimeEnv.BING_TRANSLATOR_REGION || '';

  const hasApiKey =
    normalizedEngine === 'google' ? Boolean(googleApiKey.trim()) : Boolean(bingApiKey.trim());

  const cacheKey = buildCacheKey({
    engine: normalizedEngine,
    from,
    to,
    text: rawText,
    hasApiKey
  });

  if (translationCache.has(cacheKey)) {
    return { ...translationCache.get(cacheKey), cached: true };
  }

  const result =
    normalizedEngine === 'google'
      ? await translateWithGoogle({
          text: rawText,
          from,
          to,
          apiKey: googleApiKey,
          customEndpoint: googleEndpoint,
          fetchImpl
        })
      : await translateWithBing({
          text: rawText,
          from,
          to,
          apiKey: bingApiKey,
          region: bingRegion,
          fetchImpl
        });

  setCache(cacheKey, result);
  return { ...result, cached: false };
}

/**
 * Translate an array of texts with bounded concurrency.
 */
export async function translateBatch({
  texts,
  engine = 'bing',
  from = 'auto',
  to = 'zh-CN',
  config = {},
  concurrency = 4,
  fetchImpl = fetch
}) {
  if (!Array.isArray(texts)) {
    throw new Error('texts must be an array');
  }

  const results = new Array(texts.length);
  let cursor = 0;

  async function worker() {
    while (cursor < texts.length) {
      const idx = cursor;
      cursor += 1;
      const text = texts[idx];
      try {
        const res = await translateText({ text, engine, from, to, config, fetchImpl });
        results[idx] = { ok: true, ...res };
      } catch (error) {
        results[idx] = {
          ok: false,
          engine,
          error: error.message || 'Translation failed'
        };
      }
    }
  }

  const workerCount = Math.min(Math.max(1, concurrency), Math.max(1, texts.length));
  await Promise.all(Array.from({ length: workerCount }, () => worker()));
  return results;
}
