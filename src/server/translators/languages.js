/**
 * Supported language definitions and engine-specific code mapping
 * between standard codes, Google Translate codes, and Bing/Azure Translator codes.
 */

export const SUPPORTED_LANGUAGES = [
  { code: 'auto', label: '自动检测 (Auto)', google: 'auto', bing: 'auto-detect' },
  { code: 'zh-CN', label: '简体中文 (zh-CN)', google: 'zh-CN', bing: 'zh-Hans' },
  { code: 'zh-TW', label: '繁體中文 (zh-TW)', google: 'zh-TW', bing: 'zh-Hant' },
  { code: 'en', label: '英语 (English)', google: 'en', bing: 'en' },
  { code: 'ja', label: '日语 (日本語)', google: 'ja', bing: 'ja' },
  { code: 'ko', label: '韩语 (한국어)', google: 'ko', bing: 'ko' },
  { code: 'fr', label: '法语 (Français)', google: 'fr', bing: 'fr' },
  { code: 'de', label: '德语 (Deutsch)', google: 'de', bing: 'de' },
  { code: 'es', label: '西班牙语 (Español)', google: 'es', bing: 'es' },
  { code: 'ru', label: '俄语 (Русский)', google: 'ru', bing: 'ru' },
  { code: 'pt', label: '葡萄牙语 (Português)', google: 'pt', bing: 'pt' },
  { code: 'it', label: '意大利语 (Italiano)', google: 'it', bing: 'it' },
  { code: 'ar', label: '阿拉伯语 (العربية)', google: 'ar', bing: 'ar' },
  { code: 'th', label: '泰语 (ไทย)', google: 'th', bing: 'th' },
  { code: 'vi', label: '越南语 (Tiếng Việt)', google: 'vi', bing: 'vi' }
];

const LANGUAGE_MAP = new Map(SUPPORTED_LANGUAGES.map((item) => [item.code.toLowerCase(), item]));

/**
 * Normalize a user-supplied language code to the target engine's format.
 * @param {string} code - Standard or custom language code
 * @param {'google'|'bing'} engine - Target translation engine
 * @param {boolean} isSource - Whether this is a source language (supports auto-detect)
 * @returns {string} Engine-specific language code
 */
export function mapLanguageCode(code, engine, isSource = false) {
  if (!code || code === 'auto' || code === 'auto-detect') {
    if (!isSource) {
      return engine === 'bing' ? 'zh-Hans' : 'zh-CN';
    }
    return engine === 'bing' ? 'auto-detect' : 'auto';
  }

  const normalized = code.trim();
  const entry = LANGUAGE_MAP.get(normalized.toLowerCase());
  if (entry) {
    return entry[engine];
  }

  // Handle common aliases not directly matching our canonical keys
  const lower = normalized.toLowerCase();
  if (lower === 'zh' || lower === 'zh_cn' || lower === 'zh-hans') {
    return engine === 'bing' ? 'zh-Hans' : 'zh-CN';
  }
  if (lower === 'zh_tw' || lower === 'zh-hant') {
    return engine === 'bing' ? 'zh-Hant' : 'zh-TW';
  }

  return normalized;
}
