// AES-256-GCM encryption for Google App Passwords at rest.
// Key: APP_ENCRYPTION_KEY = 32 random bytes, base64. Generate with:
//   node -e "console.log(require('crypto').randomBytes(32).toString('base64'))"
// AAD binds each ciphertext to its sender account id, so a ciphertext copied onto
// another row fails to decrypt. key_version allows future rotation
// (APP_ENCRYPTION_KEY_V2 ...) without re-entering every password at once.
import crypto from 'node:crypto';

export const CURRENT_KEY_VERSION = 1;

function keyFor(version: number): Buffer {
  const name = version === 1 ? 'APP_ENCRYPTION_KEY' : `APP_ENCRYPTION_KEY_V${version}`;
  const raw = process.env[name];
  if (!raw) throw new Error(`${name} is not set`);
  const key = Buffer.from(raw, 'base64');
  if (key.length !== 32) throw new Error(`${name} must be 32 bytes (base64-encoded)`);
  return key;
}

export interface Sealed { enc: string; iv: string; tag: string; keyVersion: number }

export function seal(plaintext: string, aad: string): Sealed {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv('aes-256-gcm', keyFor(CURRENT_KEY_VERSION), iv);
  cipher.setAAD(Buffer.from(aad, 'utf8'));
  const enc = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  return { enc: enc.toString('base64'), iv: iv.toString('base64'), tag: cipher.getAuthTag().toString('base64'), keyVersion: CURRENT_KEY_VERSION };
}

export function open(sealed: Sealed, aad: string): string {
  const decipher = crypto.createDecipheriv('aes-256-gcm', keyFor(sealed.keyVersion), Buffer.from(sealed.iv, 'base64'));
  decipher.setAAD(Buffer.from(aad, 'utf8'));
  decipher.setAuthTag(Buffer.from(sealed.tag, 'base64'));
  return Buffer.concat([decipher.update(Buffer.from(sealed.enc, 'base64')), decipher.final()]).toString('utf8');
}

export function randomToken(bytes = 32): string {
  return crypto.randomBytes(bytes).toString('base64url');
}

export function sha256Hex(s: string): string {
  return crypto.createHash('sha256').update(s).digest('hex');
}

/** HMAC-signed compact token (used for unsubscribe links). */
export function signToken(payload: object): string {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const mac = crypto.createHmac('sha256', keyFor(CURRENT_KEY_VERSION)).update('tok:' + body).digest('base64url').slice(0, 32);
  return `${body}.${mac}`;
}

export function verifyToken<T = unknown>(token: string): T | null {
  const [body, mac] = token.split('.');
  if (!body || !mac) return null;
  const expected = crypto.createHmac('sha256', keyFor(CURRENT_KEY_VERSION)).update('tok:' + body).digest('base64url').slice(0, 32);
  const a = Buffer.from(mac); const b = Buffer.from(expected);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  try { return JSON.parse(Buffer.from(body, 'base64url').toString('utf8')) as T; } catch { return null; }
}

/** Human-friendly temporary password, e.g. "Kxq7-mPz3-Wd9t". Avoids ambiguous characters. */
export function tempPassword(): string {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const pick = () => Array.from(crypto.randomBytes(4), b => alphabet[b % alphabet.length]).join('');
  return `${pick()}-${pick()}-${pick()}`;
}
