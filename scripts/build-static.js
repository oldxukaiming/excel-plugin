import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, '..');
const CLIENT_DIR = path.join(ROOT_DIR, 'src/client');
const OUT_DIR = path.join(ROOT_DIR, 'dist/static');

function parseBaseUrlArg() {
  const arg = process.argv.find((a) => a.startsWith('--url='));
  const raw = arg ? arg.slice('--url='.length) : process.env.PUBLIC_URL || '';
  return raw.trim().replace(/\/+$/, '');
}

function parseApiUrlArg() {
  const arg = process.argv.find((a) => a.startsWith('--api='));
  const raw = arg ? arg.slice('--api='.length) : process.env.API_BASE_URL || '';
  return raw.trim().replace(/\/+$/, '');
}

function main() {
  fs.rmSync(OUT_DIR, { recursive: true, force: true });
  fs.cpSync(CLIENT_DIR, OUT_DIR, { recursive: true });

  // Bundle exceljs browser build so static deployments handle .xlsx entirely client-side
  const excelJsMin = path.join(ROOT_DIR, 'node_modules/exceljs/dist/exceljs.min.js');
  if (fs.existsSync(excelJsMin)) {
    const vendorDir = path.join(OUT_DIR, 'vendor');
    fs.mkdirSync(vendorDir, { recursive: true });
    fs.copyFileSync(excelJsMin, path.join(vendorDir, 'exceljs.min.js'));
  }

  // Generate redirect index.html for static hosts (GitHub Pages / Cloudflare Pages)
  const indexHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta http-equiv="refresh" content="0; url=./workbench.html" />
  <title>Excel 双引擎翻译工作台</title>
</head>
<body>
  <p>正在跳转到 <a href="./workbench.html">Excel 双引擎翻译工作台</a>...</p>
</body>
</html>
`;
  fs.writeFileSync(path.join(OUT_DIR, 'index.html'), indexHtml, 'utf8');

  // Write runtime config.js if an external API URL is supplied (e.g. GitHub Pages -> Cloudflare Worker)
  const apiUrl = parseApiUrlArg();
  const configJs = `window.__EXCEL_TRANSLATOR_CONFIG__ = ${JSON.stringify({ apiBaseUrl: apiUrl })};\n`;
  fs.writeFileSync(path.join(OUT_DIR, 'config.js'), configJs, 'utf8');

  // Build customized manifest.xml if public URL is provided
  const publicUrl = parseBaseUrlArg();
  const manifestSrc = path.join(ROOT_DIR, 'manifest.xml');
  if (fs.existsSync(manifestSrc)) {
    let manifestXml = fs.readFileSync(manifestSrc, 'utf8');
    if (publicUrl) {
      manifestXml = manifestXml.replaceAll('https://localhost:3443', publicUrl);
    }
    fs.writeFileSync(path.join(OUT_DIR, 'manifest.xml'), manifestXml, 'utf8');
  }

  // Support GitHub Pages custom domain CNAME if provided
  const cnameArg = process.argv.find((a) => a.startsWith('--cname='));
  const cname = cnameArg ? cnameArg.slice('--cname='.length).trim() : process.env.CNAME || '';
  if (cname) {
    fs.writeFileSync(path.join(OUT_DIR, 'CNAME'), `${cname}\n`, 'utf8');
  }

  console.log(`Static bundle built at ${OUT_DIR}`);
  if (publicUrl) console.log(`  manifest.xml base URL: ${publicUrl}`);
  if (apiUrl) console.log(`  config.js apiBaseUrl: ${apiUrl}`);
}

main();
