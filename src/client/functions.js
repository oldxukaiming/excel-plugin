const STORAGE_KEY = 'excel-translator-settings-v1';

function readLocalSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed.settings || {};
  } catch {
    return {};
  }
}

/**
 * Translate text inside an Excel cell using Bing or Google engine.
 * @customfunction
 * @param {string} text Source text to translate.
 * @param {string} [targetLang="zh-CN"] Target language code (e.g. "zh-CN", "en", "ja").
 * @param {string} [sourceLang="auto"] Source language code or "auto".
 * @param {string} [engine="bing"] Translation engine: "bing" or "google".
 * @returns {Promise<string>} Translated text.
 */
async function translate(text, targetLang = 'zh-CN', sourceLang = 'auto', engine = 'bing') {
  if (text === null || text === undefined || String(text).trim() === '') {
    return '';
  }

  const normalizedEngine = String(engine || 'bing').toLowerCase();
  if (normalizedEngine !== 'bing' && normalizedEngine !== 'google') {
    throw new Error('Engine must be "bing" or "google"');
  }

  const response = await fetch('/api/translate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      text: String(text),
      engine: normalizedEngine,
      from: sourceLang || 'auto',
      to: targetLang || 'zh-CN',
      config: readLocalSettings()
    })
  });

  const data = await response.json();
  if (!response.ok || !data.ok) {
    throw new Error(data.error || 'Translation request failed');
  }

  return data.translatedText;
}

if (typeof globalThis.CustomFunctions !== 'undefined') {
  globalThis.CustomFunctions.associate('TRANSLATE', translate);
}

export { translate };
