import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, '..');
const MANIFEST_PATH = path.join(ROOT_DIR, 'manifest.xml');
const CERT_PATH = path.join(ROOT_DIR, '.certs', 'localhost.crt');
const ADDIN_ID = 'c7b6a1e4-8b2d-4f6a-9d1c-3f5a7e9b2d10';

function runCommand(cmd, args) {
  return execFileSync(cmd, args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
}

function registerManifestInCurrentUser() {
  if (process.platform !== 'win32') {
    throw new Error('HKCU registry sideloading is only supported on Windows.');
  }
  if (!fs.existsSync(MANIFEST_PATH)) {
    throw new Error(`manifest.xml not found at ${MANIFEST_PATH}`);
  }

  const keyPath = 'HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer';
  runCommand('reg', ['add', keyPath, '/v', ADDIN_ID, '/t', 'REG_SZ', '/d', MANIFEST_PATH, '/f']);
  console.log(`[OK] Registered manifest in current-user registry (no shared folder required):`);
  console.log(`     ${keyPath}\\${ADDIN_ID} -> ${MANIFEST_PATH}`);
}

function trustCertInCurrentUserStore() {
  if (process.platform !== 'win32' || !fs.existsSync(CERT_PATH)) {
    return;
  }
  try {
    runCommand('certutil', ['-user', '-addstore', 'Root', CERT_PATH]);
    console.log(`[OK] Added localhost certificate to CurrentUser\\Root store.`);
  } catch (err) {
    console.warn(`[WARN] Could not auto-add certificate to CurrentUser\\Root: ${err.message}`);
    console.warn(`       You can manually import ${CERT_PATH} into Current User -> Trusted Root Certification Authorities.`);
  }
}

function main() {
  registerManifestInCurrentUser();
  if (!process.argv.includes('--skip-cert')) {
    trustCertInCurrentUserStore();
  }
  console.log('\nNext steps:');
  console.log('1. Keep the local server running: npm start');
  console.log('2. Restart Excel, then go to Insert/Home -> Add-ins -> My Add-ins -> Developer Add-ins.');
}

try {
  main();
} catch (err) {
  console.error(`[ERROR] Registration failed: ${err.message}`);
  console.error('If HKCU registry or Office Add-ins are blocked by group policy, use the VBA fallback in src/vba/ExcelTranslator.bas.');
  process.exit(1);
}
