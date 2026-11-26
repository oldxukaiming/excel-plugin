import fs from 'node:fs';
import http from 'node:http';
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import { SUPPORTED_LANGUAGES } from './translators/languages.js';
import { translateBatch, translateText } from './translators/service.js';
import {
  computeTargetAddress,
  readLiveExcelRange,
  writeLiveExcelRange
} from './excel-live.js';
import { inspectXlsxBuffer, translateXlsxBuffer } from './xlsx-translator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, '../..');
const CLIENT_DIR = path.join(ROOT_DIR, 'src/client');

export function createApp() {
  const app = express();

  app.use(express.json({ limit: '25mb' }));

  // Allow Office Add-in task pane and custom functions runtime access
  app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') {
      return res.sendStatus(204);
    }
    return next();
  });

  app.get('/api/health', (_req, res) => {
    res.json({ ok: true, service: 'excel-translator-addin', timestamp: new Date().toISOString() });
  });

  app.get('/api/languages', (_req, res) => {
    res.json({ languages: SUPPORTED_LANGUAGES });
  });

  app.post('/api/translate', async (req, res) => {
    try {
      const { text, engine = 'bing', from = 'auto', to = 'zh-CN', config = {} } = req.body || {};
      const result = await translateText({ text, engine, from, to, config });
      res.json({ ok: true, ...result });
    } catch (error) {
      res.status(502).json({
        ok: false,
        error: error.message || 'Translation failed'
      });
    }
  });

  app.post('/api/translate/batch', async (req, res) => {
    try {
      const { texts, engine = 'bing', from = 'auto', to = 'zh-CN', config = {} } = req.body || {};
      if (!Array.isArray(texts)) {
        return res.status(400).json({ ok: false, error: 'texts must be an array' });
      }
      const results = await translateBatch({ texts, engine, from, to, config });
      return res.json({ ok: true, results });
    } catch (error) {
      return res.status(502).json({
        ok: false,
        error: error.message || 'Batch translation failed'
      });
    }
  });

  app.post('/api/xlsx/inspect', async (req, res) => {
    try {
      const { fileBase64 } = req.body || {};
      if (!fileBase64 || typeof fileBase64 !== 'string') {
        return res.status(400).json({ ok: false, error: 'fileBase64 is required' });
      }
      const buffer = Buffer.from(fileBase64, 'base64');
      const info = await inspectXlsxBuffer(buffer);
      return res.json({ ok: true, ...info });
    } catch (error) {
      return res.status(400).json({ ok: false, error: error.message || 'Failed to inspect .xlsx file' });
    }
  });

  app.post('/api/xlsx/translate', async (req, res) => {
    try {
      const {
        fileBase64,
        sheetName = '',
        sourceAddress = 'A1',
        targetMode = 'below',
        customTargetAddress = '',
        engine = 'bing',
        from = 'auto',
        to = 'zh-CN',
        config = {}
      } = req.body || {};

      if (!fileBase64 || typeof fileBase64 !== 'string') {
        return res.status(400).json({ ok: false, error: 'fileBase64 is required' });
      }

      const buffer = Buffer.from(fileBase64, 'base64');
      const result = await translateXlsxBuffer({
        buffer,
        sheetName,
        sourceAddress,
        targetMode,
        customTargetAddress,
        engine,
        from,
        to,
        config
      });

      return res.json({
        ok: true,
        fileBase64: result.outputBuffer.toString('base64'),
        sourceSheet: result.sourceSheet,
        targetSheet: result.targetSheet,
        sourceRange: result.sourceRange,
        targetRange: result.targetRange,
        translatedCount: result.translatedCount,
        failedCount: result.failedCount,
        firstError: result.firstError
      });
    } catch (error) {
      return res.status(400).json({ ok: false, error: error.message || 'Failed to translate .xlsx file' });
    }
  });

  app.get('/api/excel/selection', async (req, res) => {
    try {
      const address = typeof req.query.address === 'string' ? req.query.address : '';
      const selection = await readLiveExcelRange(address);
      return res.json(selection);
    } catch (error) {
      return res.status(502).json({ ok: false, error: error.message || 'Failed to read active Excel' });
    }
  });

  app.post('/api/excel/translate-fill', async (req, res) => {
    try {
      const {
        sourceAddress = '',
        targetMode = 'below',
        customTargetAddress = '',
        engine = 'bing',
        from = 'auto',
        to = 'zh-CN',
        config = {}
      } = req.body || {};

      const selection = await readLiveExcelRange(sourceAddress);
      const effectiveSource = sourceAddress.trim() || selection.address;
      const targetAddress = computeTargetAddress(effectiveSource, targetMode, customTargetAddress);

      const values2D = selection.values;
      const rowCount = values2D.length;
      const colCount = values2D[0]?.length || 0;
      const flat = [];
      const nonEmptyIndices = [];

      for (let r = 0; r < rowCount; r += 1) {
        for (let c = 0; c < colCount; c += 1) {
          const str = values2D[r][c] === null || values2D[r][c] === undefined ? '' : String(values2D[r][c]);
          if (str.trim() !== '') {
            nonEmptyIndices.push(flat.length);
            flat.push(str);
          } else {
            flat.push('');
          }
        }
      }

      if (nonEmptyIndices.length === 0) {
        return res.status(400).json({ ok: false, error: `No translatable text in ${effectiveSource}` });
      }

      const batch = await translateBatch({
        texts: nonEmptyIndices.map((idx) => flat[idx]),
        engine,
        from,
        to,
        config
      });

      const outFlat = [...flat];
      const errors = [];
      batch.forEach((item, idx) => {
        const flatIdx = nonEmptyIndices[idx];
        if (item.ok) {
          outFlat[flatIdx] = item.translatedText;
        } else {
          errors.push(item.error);
        }
      });

      if (errors.length === batch.length) {
        return res.status(502).json({ ok: false, error: errors[0] || 'Translation failed' });
      }

      const out2D = [];
      for (let r = 0; r < rowCount; r += 1) {
        const row = [];
        for (let c = 0; c < colCount; c += 1) {
          row.push(outFlat[r * colCount + c]);
        }
        out2D.push(row);
      }

      await writeLiveExcelRange(targetAddress, out2D);

      return res.json({
        ok: true,
        workbook: selection.workbook,
        sheet: selection.sheet,
        sourceAddress: effectiveSource,
        targetAddress,
        translatedCount: nonEmptyIndices.length - errors.length,
        failedCount: errors.length,
        preview: out2D[0]?.[0] || ''
      });
    } catch (error) {
      return res.status(502).json({ ok: false, error: error.message || 'Live Excel translation failed' });
    }
  });

  app.get('/', (_req, res) => {
    res.redirect('/workbench.html');
  });

  app.use(express.static(CLIENT_DIR));

  return app;
}

/**
 * Locate HTTPS certificates from project .certs/ directory or office-addin-dev-certs default path.
 */
export function resolveHttpsCredentials() {
  const localCertDir = path.join(ROOT_DIR, '.certs');
  const officeCertDir = path.join(os.homedir(), '.office-addin-dev-certs');

  const candidates = [
    {
      key: path.join(localCertDir, 'localhost.key'),
      cert: path.join(localCertDir, 'localhost.crt')
    },
    {
      key: path.join(officeCertDir, 'localhost.key'),
      cert: path.join(officeCertDir, 'localhost.crt')
    }
  ];

  for (const pair of candidates) {
    if (fs.existsSync(pair.key) && fs.existsSync(pair.cert)) {
      return {
        key: fs.readFileSync(pair.key),
        cert: fs.readFileSync(pair.cert),
        certPath: pair.cert
      };
    }
  }

  return null;
}

function attachPortErrorHandler(server, label, port) {
  server.on('error', (err) => {
    if (err && err.code === 'EADDRINUSE') {
      console.error(
        `[ERROR] ${label} port ${port} is already in use. The server may already be running in another terminal.`
      );
      process.exitCode = 1;
      return;
    }
    console.error(`[ERROR] ${label} server error:`, err.message);
    process.exitCode = 1;
  });
}

export function startServer(port = Number(process.env.PORT) || 3443) {
  const app = createApp();
  const forceHttp = process.argv.includes('--http') || process.env.FORCE_HTTP === '1';
  const creds = forceHttp ? null : resolveHttpsCredentials();

  if (creds) {
    const httpsServer = https.createServer({ key: creds.key, cert: creds.cert }, app);
    attachPortErrorHandler(httpsServer, 'HTTPS', port);
    httpsServer.listen(port, 'localhost', () => {
      console.log(`Excel Translator Add-in HTTPS server running at https://localhost:${port}`);
      console.log(`Using certificate: ${creds.certPath}`);
    });

    const companionHttpPort =
      process.env.HTTP_PORT !== undefined ? Number(process.env.HTTP_PORT) : port + 1;
    if (companionHttpPort > 0) {
      const httpServer = http.createServer(app);
      httpServer.on('error', (err) => {
        if (err && err.code === 'EADDRINUSE') {
          console.warn(
            `[WARN] Companion HTTP port ${companionHttpPort} is already in use; skipping companion listener.`
          );
          return;
        }
        console.warn(`[WARN] Companion HTTP server error: ${err.message}`);
      });
      httpServer.listen(companionHttpPort, '127.0.0.1', () => {
        console.log(
          `Companion HTTP server (for VBA / local tools) running at http://127.0.0.1:${companionHttpPort}`
        );
      });
    }

    return httpsServer;
  }

  const fallbackPort = Number(process.env.HTTP_PORT) || port;
  const server = http.createServer(app);
  attachPortErrorHandler(server, 'HTTP', fallbackPort);
  server.listen(fallbackPort, 'localhost', () => {
    console.warn(
      `[WARN] HTTPS certificate not found. Running HTTP server at http://localhost:${fallbackPort}`
    );
    console.warn('Run "npm run cert:gen" to generate trusted localhost HTTPS certificates for Excel.');
  });
  return server;
}

if (process.argv[1] && path.resolve(process.argv[1]) === __filename) {
  startServer();
}
