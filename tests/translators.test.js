import { beforeEach, describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { mapLanguageCode } from '../src/server/translators/languages.js';
import { translateWithGoogle } from '../src/server/translators/google.js';
import {
  extractCookieHeader,
  parseBingTranslatorHtml,
  resetBingAuthCache,
  translateWithBing
} from '../src/server/translators/bing.js';
import {
  clearTranslationCache,
  translateBatch,
  translateText
} from '../src/server/translators/service.js';

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' }
  });
}

function textResponse(body, status = 200, extraHeaders = {}) {
  return new Response(body, {
    status,
    headers: { 'Content-Type': 'text/plain', ...extraHeaders }
  });
}

describe('translators', () => {
  beforeEach(() => {
    resetBingAuthCache();
    clearTranslationCache();
  });

  it('maps language codes accurately between Google and Bing', () => {
    assert.equal(mapLanguageCode('auto', 'google', true), 'auto');
    assert.equal(mapLanguageCode('auto', 'bing', true), 'auto-detect');
    assert.equal(mapLanguageCode('zh-CN', 'google', false), 'zh-CN');
    assert.equal(mapLanguageCode('zh-CN', 'bing', false), 'zh-Hans');
    assert.equal(mapLanguageCode('zh-TW', 'bing', false), 'zh-Hant');
  });

  it('translates with Google clients5 primary endpoint (auto & explicit source)', async () => {
    const mockFetchAuto = async (url) => {
      assert.match(url, /clients5\.google\.com\/translate_a\/t/);
      return jsonResponse([['你好，Excel', 'en']]);
    };

    const resAuto = await translateWithGoogle({
      text: 'Hello, Excel',
      from: 'auto',
      to: 'zh-CN',
      fetchImpl: mockFetchAuto
    });
    assert.equal(resAuto.engine, 'google');
    assert.equal(resAuto.translatedText, '你好，Excel');
    assert.equal(resAuto.detectedSourceLanguage, 'en');

    const mockFetchExplicit = async (url) => {
      assert.match(url, /clients5\.google\.com\/translate_a\/t/);
      return jsonResponse(['你好世界']);
    };

    const resExplicit = await translateWithGoogle({
      text: 'Hello world',
      from: 'en',
      to: 'zh-CN',
      fetchImpl: mockFetchExplicit
    });
    assert.equal(resExplicit.translatedText, '你好世界');
  });

  it('falls back to translate_a/single and joins segments when clients5 fails', async () => {
    const mockFetch = async (url) => {
      if (url.includes('clients5.google.com')) {
        return textResponse('Too Many Requests', 429);
      }
      assert.match(url, /translate_a\/single/);
      return jsonResponse([
        [
          ['你好，', 'Hello, '],
          ['世界', 'world']
        ],
        null,
        'en'
      ]);
    };

    const res = await translateWithGoogle({
      text: 'Hello, world',
      from: 'auto',
      to: 'zh-CN',
      fetchImpl: mockFetch
    });

    assert.equal(res.engine, 'google');
    assert.equal(res.mode, 'free-web');
    assert.equal(res.translatedText, '你好，世界');
    assert.equal(res.detectedSourceLanguage, 'en');
  });

  it('translates with Google official API when apiKey is provided', async () => {
    const mockFetch = async (url, init) => {
      assert.match(url, /translation\.googleapis\.com\/language\/translate\/v2\?key=test-g-key/);
      const body = JSON.parse(init.body);
      assert.equal(body.q, 'Hello');
      assert.equal(body.target, 'zh-CN');
      return jsonResponse({
        data: {
          translations: [{ translatedText: '你好', detectedSourceLanguage: 'en' }]
        }
      });
    };

    const res = await translateWithGoogle({
      text: 'Hello',
      from: 'auto',
      to: 'zh-CN',
      apiKey: 'test-g-key',
      fetchImpl: mockFetch
    });

    assert.equal(res.mode, 'official-api');
    assert.equal(res.translatedText, '你好');
  });

  it('translates with Bing web endpoint, extracts cookies, and forwards them', async () => {
    const sampleHtml = `
      <html>
        <script>
          IG:"ABCDEF123456";
          var params_AbusePreventionHelper = [1732599999999,"web-token-xyz",1800000];
        </script>
        <div data-iid="translator.5024"></div>
      </html>
    `;

    const parsed = parseBingTranslatorHtml(sampleHtml);
    assert.equal(parsed.ig, 'ABCDEF123456');
    assert.equal(parsed.token, 'web-token-xyz');

    const headersWithCookie = new Headers();
    headersWithCookie.append('set-cookie', 'MUID=12345; path=/; Secure');
    headersWithCookie.append('set-cookie', 'SRCHD=AF=NOFORM; path=/');
    assert.equal(extractCookieHeader(headersWithCookie), 'MUID=12345; SRCHD=AF=NOFORM');

    let sawCookie = '';
    const mockFetch = async (url, init) => {
      if (url.includes('/translator')) {
        const resp = new Response(sampleHtml, { status: 200, headers: headersWithCookie });
        Object.defineProperty(resp, 'url', { value: 'https://cn.bing.com/translator' });
        return resp;
      }
      if (url.includes('/ttranslatev3')) {
        assert.match(url, /^https:\/\/cn\.bing\.com\/ttranslatev3/);
        sawCookie = init.headers.Cookie || '';
        return jsonResponse([
          {
            detectedLanguage: { language: 'en' },
            translations: [{ text: '必应网页翻译', to: 'zh-Hans' }]
          }
        ]);
      }
      throw new Error(`Unexpected URL: ${url}`);
    };

    const res = await translateWithBing({
      text: 'Bing Web Translate',
      from: 'auto',
      to: 'zh-CN',
      fetchImpl: mockFetch
    });

    assert.equal(res.mode, 'free-web');
    assert.equal(res.translatedText, '必应网页翻译');
    assert.equal(sawCookie, 'MUID=12345; SRCHD=AF=NOFORM');
  });

  it('falls back to Bing Edge endpoint when web endpoint fails', async () => {
    let authCalls = 0;
    let translateCalls = 0;

    const mockFetch = async (url) => {
      if (url.includes('/translator')) {
        return textResponse('Blocked', 403);
      }
      if (url.includes('edge.microsoft.com/translate/auth')) {
        authCalls += 1;
        return textResponse(`jwt-token-${authCalls}`);
      }
      if (url.includes('api-edge.cognitive.microsofttranslator.com/translate')) {
        translateCalls += 1;
        if (translateCalls === 1) {
          return textResponse('Unauthorized', 401);
        }
        return jsonResponse([
          {
            detectedLanguage: { language: 'en' },
            translations: [{ text: 'Excel 翻译', to: 'zh-Hans' }]
          }
        ]);
      }
      throw new Error(`Unexpected URL: ${url}`);
    };

    const res = await translateWithBing({
      text: 'Excel Translation',
      from: 'auto',
      to: 'zh-CN',
      fetchImpl: mockFetch
    });

    assert.equal(authCalls, 2);
    assert.equal(translateCalls, 2);
    assert.equal(res.engine, 'bing');
    assert.equal(res.mode, 'free-edge');
    assert.equal(res.translatedText, 'Excel 翻译');
  });

  it('switches engines and caches results per engine in translateText & translateBatch', async () => {
    const sampleHtml = `
      <script>
        IG:"IG123";
        var params_AbusePreventionHelper = [1732599999999,"tok",1800000];
      </script>
      <div data-iid="translator.5024"></div>
    `;
    let calls = 0;
    const mockFetch = async (url) => {
      calls += 1;
      if (url.includes('/translator')) {
        return textResponse(sampleHtml);
      }
      if (url.includes('/ttranslatev3')) {
        return jsonResponse([{ translations: [{ text: '来自必应' }] }]);
      }
      if (url.includes('clients5.google.com')) {
        return jsonResponse([['来自谷歌', 'en']]);
      }
      throw new Error(`Unexpected URL: ${url}`);
    };

    const bingRes1 = await translateText({ text: 'Hello', engine: 'bing', fetchImpl: mockFetch });
    const bingRes2 = await translateText({ text: 'Hello', engine: 'bing', fetchImpl: mockFetch });
    const googleRes = await translateText({ text: 'Hello', engine: 'google', fetchImpl: mockFetch });

    assert.equal(bingRes1.translatedText, '来自必应');
    assert.equal(bingRes1.cached, false);
    assert.equal(bingRes2.cached, true);
    assert.equal(googleRes.translatedText, '来自谷歌');
    assert.equal(googleRes.cached, false);
    // 1 bing html + 1 bing translate + 1 google translate = 3 network calls
    assert.equal(calls, 3);

    const batch = await translateBatch({
      texts: ['Hello', '   ', 'Hello'],
      engine: 'google',
      fetchImpl: mockFetch
    });
    assert.equal(batch.length, 3);
    assert.equal(batch[0].translatedText, '来自谷歌');
    assert.equal(batch[1].translatedText, '   ');
    assert.equal(batch[2].cached, true);
  });
});
