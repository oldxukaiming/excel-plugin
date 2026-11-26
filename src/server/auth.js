/**
 * Shared username/password authentication & HMAC-SHA256 session token module.
 * Uses standard Web Crypto API so it works identically in Cloudflare Workers and Node.js 20+.
 */

const textEncoder = new TextEncoder();

function getRuntimeEnv(env = {}) {
  const procEnv = typeof process !== 'undefined' && process?.env ? process.env : {};
  return {
    AUTH_USERNAME: env.AUTH_USERNAME ?? procEnv.AUTH_USERNAME ?? '',
    AUTH_PASSWORD: env.AUTH_PASSWORD ?? procEnv.AUTH_PASSWORD ?? '',
    AUTH_SECRET: env.AUTH_SECRET ?? procEnv.AUTH_SECRET ?? ''
  };
}

export function getAuthConfig(env = {}) {
  const resolved = getRuntimeEnv(env);
  const password = String(resolved.AUTH_PASSWORD || '');
  const enabled = password.length > 0;
  const username = String(resolved.AUTH_USERNAME || 'admin');
  const secret = String(resolved.AUTH_SECRET || `${password}:excel-translator-hmac-v1`);
  return { enabled, username, password, secret };
}

/**
 * Constant-time string comparison to avoid timing side-channels.
 */
export function timingSafeEqualStrings(a, b) {
  const bytesA = textEncoder.encode(String(a ?? ''));
  const bytesB = textEncoder.encode(String(b ?? ''));
  const maxLen = Math.max(bytesA.length, bytesB.length);
  let diff = bytesA.length ^ bytesB.length;
  for (let i = 0; i < maxLen; i += 1) {
    const va = i < bytesA.length ? bytesA[i] : 0;
    const vb = i < bytesB.length ? bytesB[i] : 0;
    diff |= va ^ vb;
  }
  return diff === 0;
}

function toBase64Url(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function fromBase64Url(str) {
  const padded = str.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((str.length + 3) % 4);
  const binary = atob(padded);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

async function importHmacKey(secret) {
  return crypto.subtle.importKey(
    'raw',
    textEncoder.encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign', 'verify']
  );
}

export async function signSessionToken(username, secret, ttlSeconds = 7 * 24 * 3600) {
  const payload = {
    u: username,
    exp: Math.floor(Date.now() / 1000) + ttlSeconds
  };
  const payloadBytes = textEncoder.encode(JSON.stringify(payload));
  const payloadB64 = toBase64Url(payloadBytes);
  const key = await importHmacKey(secret);
  const sigBuffer = await crypto.subtle.sign('HMAC', key, textEncoder.encode(payloadB64));
  const sigB64 = toBase64Url(new Uint8Array(sigBuffer));
  return `${payloadB64}.${sigB64}`;
}

export async function verifySessionToken(token, secret) {
  if (!token || typeof token !== 'string') {
    return { valid: false };
  }
  const dotIdx = token.indexOf('.');
  if (dotIdx <= 0 || dotIdx === token.length - 1) {
    return { valid: false };
  }

  const payloadB64 = token.slice(0, dotIdx);
  const sigB64 = token.slice(dotIdx + 1);

  try {
    const key = await importHmacKey(secret);
    const sigBytes = fromBase64Url(sigB64);
    const ok = await crypto.subtle.verify('HMAC', key, sigBytes, textEncoder.encode(payloadB64));
    if (!ok) return { valid: false };

    const payloadJson = new TextDecoder().decode(fromBase64Url(payloadB64));
    const payload = JSON.parse(payloadJson);
    if (!payload || typeof payload.exp !== 'number' || payload.exp < Math.floor(Date.now() / 1000)) {
      return { valid: false };
    }
    return { valid: true, username: String(payload.u || '') };
  } catch {
    return { valid: false };
  }
}

export function validateCredentials(inputUser, inputPass, config) {
  if (!config.enabled) return false;
  return (
    timingSafeEqualStrings(String(inputUser || '').trim(), config.username) &&
    timingSafeEqualStrings(String(inputPass || ''), config.password)
  );
}

function getHeader(headers, name) {
  if (!headers) return '';
  if (typeof headers.get === 'function') {
    return headers.get(name) || '';
  }
  const lower = name.toLowerCase();
  for (const [k, v] of Object.entries(headers)) {
    if (k.toLowerCase() === lower) {
      return Array.isArray(v) ? v[0] : String(v || '');
    }
  }
  return '';
}

/**
 * Authorize an incoming request via Bearer session token, X-Access-Token, or HTTP Basic auth.
 */
export async function authorizeRequest(headers, env = {}) {
  const config = getAuthConfig(env);
  if (!config.enabled) {
    return { authorized: true, authRequired: false, username: null };
  }

  const authHeader = getHeader(headers, 'authorization').trim();
  const customToken = getHeader(headers, 'x-access-token').trim();

  if (authHeader.toLowerCase().startsWith('bearer ')) {
    const token = authHeader.slice(7).trim();
    const verified = await verifySessionToken(token, config.secret);
    if (verified.valid) {
      return { authorized: true, authRequired: true, username: verified.username };
    }
  } else if (authHeader.toLowerCase().startsWith('basic ')) {
    try {
      const decoded = atob(authHeader.slice(6).trim());
      const colonIdx = decoded.indexOf(':');
      if (colonIdx !== -1) {
        const u = decoded.slice(0, colonIdx);
        const p = decoded.slice(colonIdx + 1);
        if (validateCredentials(u, p, config)) {
          return { authorized: true, authRequired: true, username: config.username };
        }
      }
    } catch {
      // Ignore malformed Basic auth header
    }
  }

  if (customToken) {
    const verified = await verifySessionToken(customToken, config.secret);
    if (verified.valid) {
      return { authorized: true, authRequired: true, username: verified.username };
    }
  }

  return { authorized: false, authRequired: true, username: null };
}
