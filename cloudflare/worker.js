import {
  authorizeRequest,
  getAuthConfig,
  signSessionToken,
  validateCredentials
} from '../src/server/auth.js';
import { SUPPORTED_LANGUAGES } from '../src/server/translators/languages.js';
import { translateBatch, translateText } from '../src/server/translators/service.js';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Access-Token'
};

function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      ...CORS_HEADERS
    }
  });
}

function mergeEnvConfig(bodyConfig = {}, env = {}) {
  return {
    googleApiKey: bodyConfig.googleApiKey || env.GOOGLE_TRANSLATE_API_KEY || '',
    googleEndpoint: bodyConfig.googleEndpoint || env.GOOGLE_TRANSLATE_ENDPOINT || '',
    bingApiKey: bodyConfig.bingApiKey || env.BING_TRANSLATOR_API_KEY || '',
    bingRegion: bodyConfig.bingRegion || env.BING_TRANSLATOR_REGION || ''
  };
}

export async function handleWorkerRequest(request, env = {}) {
  const url = new URL(request.url);
  const { pathname } = url;

  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 204, headers: CORS_HEADERS });
  }

  const authConfig = getAuthConfig(env);

  if (request.method === 'GET' && pathname === '/api/health') {
    return jsonResponse({
      ok: true,
      service: 'excel-translator-addin-worker',
      version: '1.2.0',
      runtime: 'cloudflare-workers',
      authRequired: authConfig.enabled,
      timestamp: new Date().toISOString()
    });
  }

  if (request.method === 'GET' && pathname === '/api/auth/status') {
    const auth = await authorizeRequest(request.headers, env);
    return jsonResponse({
      ok: true,
      authRequired: auth.authRequired,
      authenticated: auth.authorized,
      username: auth.username
    });
  }

  if (request.method === 'POST' && pathname === '/api/auth/login') {
    if (!authConfig.enabled) {
      return jsonResponse({
        ok: true,
        authRequired: false,
        token: '',
        username: 'guest'
      });
    }
    try {
      const body = await request.json();
      const { username = '', password = '' } = body || {};
      if (!validateCredentials(username, password, authConfig)) {
        return jsonResponse({ ok: false, error: '用户名或密码错误' }, 401);
      }
      const ttlSeconds = 7 * 24 * 3600;
      const token = await signSessionToken(authConfig.username, authConfig.secret, ttlSeconds);
      return jsonResponse({
        ok: true,
        authRequired: true,
        token,
        username: authConfig.username,
        expiresIn: ttlSeconds
      });
    } catch {
      return jsonResponse({ ok: false, error: '登录请求格式无效' }, 400);
    }
  }

  // Require valid authentication for all protected API endpoints when AUTH_PASSWORD is configured
  if (pathname.startsWith('/api/')) {
    const auth = await authorizeRequest(request.headers, env);
    if (!auth.authorized) {
      return jsonResponse(
        { ok: false, authRequired: true, error: '访问受限，请先登录' },
        401
      );
    }
  }

  if (request.method === 'GET' && pathname === '/api/languages') {
    return jsonResponse({ languages: SUPPORTED_LANGUAGES });
  }

  if (request.method === 'POST' && pathname === '/api/translate') {
    try {
      const body = await request.json();
      const { text, engine = 'bing', from = 'auto', to = 'zh-CN', config = {} } = body || {};
      const result = await translateText({
        text,
        engine,
        from,
        to,
        config: mergeEnvConfig(config, env)
      });
      return jsonResponse({ ok: true, ...result });
    } catch (error) {
      return jsonResponse({ ok: false, error: error.message || 'Translation failed' }, 502);
    }
  }

  if (request.method === 'POST' && pathname === '/api/translate/batch') {
    try {
      const body = await request.json();
      const { texts, engine = 'bing', from = 'auto', to = 'zh-CN', config = {} } = body || {};
      if (!Array.isArray(texts)) {
        return jsonResponse({ ok: false, error: 'texts must be an array' }, 400);
      }
      const results = await translateBatch({
        texts,
        engine,
        from,
        to,
        config: mergeEnvConfig(config, env)
      });
      return jsonResponse({ ok: true, results });
    } catch (error) {
      return jsonResponse({ ok: false, error: error.message || 'Batch translation failed' }, 502);
    }
  }

  // When deployed with Cloudflare Workers Static Assets, serve frontend files for non-API routes
  if (env && env.ASSETS && typeof env.ASSETS.fetch === 'function') {
    if (pathname === '/' || pathname === '') {
      return Response.redirect(new URL('/workbench.html', url.origin).toString(), 302);
    }
    const assetResponse = await env.ASSETS.fetch(request);
    if (assetResponse.status !== 404) {
      return assetResponse;
    }
  }

  return jsonResponse({ ok: false, error: `Route not found: ${pathname}` }, 404);
}

export default {
  fetch(request, env) {
    return handleWorkerRequest(request, env);
  }
};
