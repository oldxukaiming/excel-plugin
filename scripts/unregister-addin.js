import { execFileSync } from 'node:child_process';

const ADDIN_ID = 'c7b6a1e4-8b2d-4f6a-9d1c-3f5a7e9b2d10';
const keyPath = 'HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer';

if (process.platform !== 'win32') {
  console.error('Registry unregistration is only supported on Windows.');
  process.exit(1);
}

try {
  execFileSync('reg', ['delete', keyPath, '/v', ADDIN_ID, '/f'], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe']
  });
  console.log(`[OK] Removed ${ADDIN_ID} from ${keyPath}`);
} catch (err) {
  console.error(`[WARN] Could not remove registry entry: ${err.message}`);
}
