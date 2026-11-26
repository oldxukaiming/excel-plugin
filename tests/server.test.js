import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { Duplex } from 'node:stream';
import { createApp } from '../src/server/index.js';

/**
 * Invoke an Express app using real Node IncomingMessage/ServerResponse backed by
 * an in-memory Duplex stream (no TCP socket required).
 */
function dispatchRequest(app, { method = 'GET', url = '/', headers = {}, body = null }) {
  return new Promise((resolve, reject) => {
    const rawBody = body === null ? null : typeof body === 'string' ? body : JSON.stringify(body);
    const reqHeaders = { ...headers };
    if (rawBody !== null) {
      reqHeaders['content-length'] = String(Buffer.byteLength(rawBody));
      if (!reqHeaders['content-type']) {
        reqHeaders['content-type'] = 'application/json';
      }
    }

    const rawChunks = [];
    const socket = new Duplex({
      read() {
        if (rawBody !== null) {
          this.push(rawBody);
        }
        this.push(null);
      },
      write(chunk, _encoding, cb) {
        rawChunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
        cb();
      }
    });
    socket.remoteAddress = '127.0.0.1';
    socket.cork = () => {};
    socket.uncork = () => {};

    const req = new http.IncomingMessage(socket);
    req.method = method;
    req.url = url;
    req.headers = reqHeaders;
    if (rawBody !== null) {
      req.push(rawBody);
    }
    req.push(null);

    const res = new http.ServerResponse(req);
    const bodyChunks = [];
    const origWrite = res.write.bind(res);
    const origEnd = res.end.bind(res);

    res.write = (chunk, encoding, cb) => {
      if (chunk) {
        bodyChunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk, encoding));
      }
      return origWrite(chunk, encoding, cb);
    };

    res.end = (chunk, encoding, cb) => {
      if (chunk && typeof chunk !== 'function') {
        bodyChunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk, encoding));
      }
      const result = origEnd(chunk, encoding, cb);
      const buffer = Buffer.concat(bodyChunks);
      resolve({
        status: res.statusCode,
        headers: res.getHeaders(),
        buffer,
        text: () => buffer.toString('utf8'),
        json: () => JSON.parse(buffer.toString('utf8'))
      });
      return result;
    };

    app.handle(req, res, (err) => {
      if (err) {
        reject(err);
      } else if (!res.writableEnded) {
        res.statusCode = 404;
        res.end('Not Found');
      }
    });
  });
}

describe('express server API and static hosting', () => {
  const app = createApp();

  it('serves /api/health and /api/languages with CORS headers', async () => {
    const healthRes = await dispatchRequest(app, { method: 'GET', url: '/api/health' });
    assert.equal(healthRes.status, 200);
    assert.equal(healthRes.headers['access-control-allow-origin'], '*');
    assert.equal(healthRes.json().ok, true);

    const langRes = await dispatchRequest(app, { method: 'GET', url: '/api/languages' });
    assert.equal(langRes.status, 200);
    const langJson = langRes.json();
    assert.ok(Array.isArray(langJson.languages));
    assert.ok(langJson.languages.some((l) => l.code === 'zh-CN'));
  });

  it('validates /api/translate/batch payload', async () => {
    const res = await dispatchRequest(app, {
      method: 'POST',
      url: '/api/translate/batch',
      body: { texts: 'not-an-array' }
    });
    assert.equal(res.status, 400);
    assert.equal(res.json().ok, false);
  });

  it('handles /api/translate for blank text without external network', async () => {
    const res = await dispatchRequest(app, {
      method: 'POST',
      url: '/api/translate',
      body: { text: '   ', engine: 'google', from: 'auto', to: 'zh-CN' }
    });
    assert.equal(res.status, 200);
    const json = res.json();
    assert.equal(json.ok, true);
    assert.equal(json.mode, 'passthrough');
  });

  it('serves taskpane.html, workbench.html, functions.json, and generated icon assets', async () => {
    const htmlRes = await dispatchRequest(app, { method: 'GET', url: '/taskpane.html' });
    assert.equal(htmlRes.status, 200);
    assert.match(htmlRes.text(), /Excel 翻译助手/);

    const wbRes = await dispatchRequest(app, { method: 'GET', url: '/workbench.html' });
    assert.equal(wbRes.status, 200);
    assert.match(wbRes.text(), /Excel 双引擎翻译工作台/);

    const fnRes = await dispatchRequest(app, { method: 'GET', url: '/functions.json' });
    assert.equal(fnRes.status, 200);
    const fnJson = fnRes.json();
    assert.equal(fnJson.functions[0].id, 'TRANSLATE');

    const iconRes = await dispatchRequest(app, { method: 'GET', url: '/assets/icon-32.png' });
    assert.equal(iconRes.status, 200);
    assert.match(String(iconRes.headers['content-type'] || ''), /image\/png/);
  });

  it('validates /api/xlsx/inspect payload', async () => {
    const res = await dispatchRequest(app, {
      method: 'POST',
      url: '/api/xlsx/inspect',
      body: {}
    });
    assert.equal(res.status, 400);
    assert.equal(res.json().ok, false);
  });

  it('enforces Express auth when AUTH_PASSWORD is set', async () => {
    const prevUser = process.env.AUTH_USERNAME;
    const prevPass = process.env.AUTH_PASSWORD;
    process.env.AUTH_USERNAME = 'alice';
    process.env.AUTH_PASSWORD = 'secret-pass';

    try {
      const blocked = await dispatchRequest(app, { method: 'GET', url: '/api/languages' });
      assert.equal(blocked.status, 401);

      const login = await dispatchRequest(app, {
        method: 'POST',
        url: '/api/auth/login',
        body: { username: 'alice', password: 'secret-pass' }
      });
      assert.equal(login.status, 200);
      const { token } = login.json();

      const allowed = await dispatchRequest(app, {
        method: 'GET',
        url: '/api/languages',
        headers: { authorization: `Bearer ${token}` }
      });
      assert.equal(allowed.status, 200);
    } finally {
      if (prevUser === undefined) delete process.env.AUTH_USERNAME;
      else process.env.AUTH_USERNAME = prevUser;
      if (prevPass === undefined) delete process.env.AUTH_PASSWORD;
      else process.env.AUTH_PASSWORD = prevPass;
    }
  });
});
