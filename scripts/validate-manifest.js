import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, '..');
const CLIENT_DIR = path.join(ROOT_DIR, 'src/client');

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function main() {
  const manifestPath = path.join(ROOT_DIR, 'manifest.xml');
  assert(fs.existsSync(manifestPath), 'manifest.xml not found');

  const xml = fs.readFileSync(manifestPath, 'utf8');

  const idMatch = xml.match(/<Id>([^<]+)<\/Id>/);
  assert(idMatch, 'Missing <Id> in manifest.xml');
  const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  assert(guidRegex.test(idMatch[1].trim()), `Invalid GUID in <Id>: ${idMatch[1]}`);

  assert(xml.includes('<Host Name="Workbook" />'), 'Manifest must target Workbook host');
  assert(xml.includes('VersionOverridesV1_0'), 'Manifest missing VersionOverridesV1_0');

  // Verify every https://localhost:3443/... referenced file exists in src/client
  const urlMatches = [...xml.matchAll(/https:\/\/localhost:3443\/([^\s"'<]+)/g)];
  assert(urlMatches.length > 0, 'No localhost:3443 resource URLs found in manifest.xml');

  const checked = new Set();
  for (const match of urlMatches) {
    const relPath = match[1];
    if (checked.has(relPath)) continue;
    checked.add(relPath);
    const fullPath = path.join(CLIENT_DIR, relPath);
    assert(fs.existsSync(fullPath), `Referenced manifest resource missing: ${fullPath}`);
  }

  // Verify functions.json is valid JSON
  const fnMetaPath = path.join(CLIENT_DIR, 'functions.json');
  const fnMeta = JSON.parse(fs.readFileSync(fnMetaPath, 'utf8'));
  assert(Array.isArray(fnMeta.functions) && fnMeta.functions.length > 0, 'functions.json must define at least one function');

  console.log(`Manifest validation passed (${checked.size} referenced client assets verified).`);
}

try {
  main();
} catch (err) {
  console.error(`Manifest validation failed: ${err.message}`);
  process.exit(1);
}
