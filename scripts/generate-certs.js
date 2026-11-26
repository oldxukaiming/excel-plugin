import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import devCerts from 'office-addin-dev-certs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, '..');
const LOCAL_CERT_DIR = path.join(ROOT_DIR, '.certs');

async function main() {
  const shouldInstallCa = process.argv.includes('--install');

  if (shouldInstallCa) {
    console.log('Ensuring localhost HTTPS certificates are generated and trusted...');
    await devCerts.ensureCertificatesAreInstalled();
  } else {
    console.log('Generating localhost HTTPS certificate files (pass --install to trust CA)...');
    await devCerts.generateCertificates();
  }

  const officeCertDir = path.join(os.homedir(), '.office-addin-dev-certs');
  const srcKey = path.join(officeCertDir, 'localhost.key');
  const srcCert = path.join(officeCertDir, 'localhost.crt');

  if (!fs.existsSync(srcKey) || !fs.existsSync(srcCert)) {
    throw new Error(`Expected certificates not found in ${officeCertDir}`);
  }

  fs.mkdirSync(LOCAL_CERT_DIR, { recursive: true });
  fs.copyFileSync(srcKey, path.join(LOCAL_CERT_DIR, 'localhost.key'));
  fs.copyFileSync(srcCert, path.join(LOCAL_CERT_DIR, 'localhost.crt'));

  console.log(`Certificates ready at ${LOCAL_CERT_DIR}`);
}

main().catch((err) => {
  console.error('Failed to generate certificates:', err.message);
  process.exit(1);
});
