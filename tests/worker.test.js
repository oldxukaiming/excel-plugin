import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { handleWorkerRequest } from '../cloudflare/worker.js';

describe('cloudflare worker handler', () => {
  it('handles CORS preflight and /api/health', async () => {
    const preflight = await handleWorkerRequest(
      new Request('https://example.com/api/translate', { method: 'OPTIONS' })
    );
    assert.equal(preflight.status, 204);
    assert.equal(preflight.headers.get('access-control-allow-origin'), '*');

    const health = await handleWorkerRequest(
      new Request('https://example.com/api/health', { method: 'GET' })
    );
    assert.equal(health.status, 200);
    const json = await health.json();
    assert.equal(json.ok, true);
    assert.equal(json.runtime, 'cloudflare-workers');
  });

  it('serves /api/languages and validates /api/translate/batch', async () => {
    const langRes = await handleWorkerRequest(
      new Request('https://example.com/api/languages', { method: 'GET' })
    );
    assert.equal(langRes.status, 200);
    const langJson = await langRes.json();
    assert.ok(langJson.languages.some((l) => l.code === 'zh-CN'));

    const badBatch = await handleWorkerRequest(
      new Request('https://example.com/api/translate/batch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ texts: 'invalid' })
      })
    );
    assert.equal(badBatch.status, 400);
    assert.equal((await badBatch.json()).ok, false);
  });

  it('delegates non-API requests to ASSETS binding when available', async () => {
    const fakeAssets = {
      async fetch(req) {
        const u = new URL(req.url);
        if (u.pathname === '/workbench.html') {
          return new Response('<html>Workbench</html>', { status: 200 });
        }
        return new Response('Not Found', { status: 404 });
      }
    };

    const rootRes = await handleWorkerRequest(new Request('https://example.com/'), {
      ASSETS: fakeAssets
    });
    assert.equal(rootRes.status, 302);
    assert.equal(rootRes.headers.get('location'), 'https://example.com/workbench.html');

    const pageRes = await handleWorkerRequest(new Request('https://example.com/workbench.html'), {
      ASSETS: fakeAssets
    });
    assert.equal(pageRes.status, 200);
    assert.equal(await pageRes.text(), '<html>Workbench</html>');
  });
});
