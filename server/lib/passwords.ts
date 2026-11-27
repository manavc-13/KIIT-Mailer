// User password hashing and the password policy.
//
// Hashing uses scrypt from Node's built-in crypto module (memory-hard, OWASP-recommended
// parameters N = 2^15, r = 8, p = 1, ~32 MiB). Being built in, it needs no native add-on,
// so the API bundles to plain JavaScript and runs identically on Windows dev machines and
// Vercel's Linux runtime.
// Stored format: $scrypt$ln=15,r=8,p=1$<salt base64>$<hash base64>
import crypto from 'node:crypto';
import { promisify } from 'node:util';

const scrypt = promisify(crypto.scrypt) as (pw: crypto.BinaryLike, salt: crypto.BinaryLike, keylen: number, opts: crypto.ScryptOptions) => Promise<Buffer>;

const LOG_N = 15;
const R = 8;
const P = 1;
const KEY_LEN = 32;
const MAXMEM = 64 * 1024 * 1024;

export const PASSWORD_MIN_LENGTH = 10;
export const PASSWORD_MAX_LENGTH = 256;

export async function hashPassword(password: string): Promise<string> {
  const salt = crypto.randomBytes(16);
  const key = await scrypt(password.normalize('NFKC'), salt, KEY_LEN, { N: 2 ** LOG_N, r: R, p: P, maxmem: MAXMEM });
  return `$scrypt$ln=${LOG_N},r=${R},p=${P}$${salt.toString('base64')}$${key.toString('base64')}`;
}

/** Constant-time verify. Returns false (never throws) on a malformed or unknown hash format. */
export async function verifyPassword(passwordHash: string, password: string): Promise<boolean> {
  try {
    const m = /^\$scrypt\$ln=(\d{1,2}),r=(\d{1,2}),p=(\d{1,2})\$([A-Za-z0-9+/=]+)\$([A-Za-z0-9+/=]+)$/.exec(passwordHash);
    if (!m) return false;
    const [, ln, r, p, saltB64, keyB64] = m;
    const expected = Buffer.from(keyB64, 'base64');
    const key = await scrypt(password.normalize('NFKC'), Buffer.from(saltB64, 'base64'), expected.length, {
      N: 2 ** Number(ln), r: Number(r), p: Number(p), maxmem: MAXMEM,
    });
    return key.length === expected.length && crypto.timingSafeEqual(key, expected);
  } catch {
    return false;
  }
}

let dummyHash: Promise<string> | null = null;
/**
 * Burn the same CPU/memory as a real verify when the email is unknown or the
 * account cannot log in, so response timing does not reveal which emails exist.
 */
export async function verifyDummy(password: string): Promise<void> {
  dummyHash ??= hashPassword('kiit-mailer-dummy-password-for-timing');
  await verifyPassword(await dummyHash, password);
}

/** Character classes present in the password (lower, upper, digit, other). */
function classCount(pw: string): number {
  let n = 0;
  if (/[a-z]/.test(pw)) n++;
  if (/[A-Z]/.test(pw)) n++;
  if (/[0-9]/.test(pw)) n++;
  if (/[^a-zA-Z0-9]/.test(pw)) n++;
  return n;
}

export interface PolicyCheck { id: 'length' | 'email' | 'variety'; label: string; ok: boolean }

/** Every rule with its pass state; the client mirrors these labels for live hints. */
export function passwordPolicyChecks(pw: string, email: string): PolicyCheck[] {
  const local = email.split('@')[0]?.trim().toLowerCase() ?? '';
  const lower = pw.toLowerCase();
  return [
    { id: 'length', label: `At least ${PASSWORD_MIN_LENGTH} characters`, ok: pw.length >= PASSWORD_MIN_LENGTH && pw.length <= PASSWORD_MAX_LENGTH },
    { id: 'email', label: 'Does not contain your email name', ok: !(local.length >= 3 && lower.includes(local)) && lower !== email.trim().toLowerCase() },
    { id: 'variety', label: 'Mixes at least two of: lowercase, uppercase, digits, symbols', ok: classCount(pw) >= 2 },
  ];
}

/** Returns a plain-English error message, or null when the password is acceptable. */
export function validatePasswordPolicy(pw: string, email: string): string | null {
  if (pw.length < PASSWORD_MIN_LENGTH) return `Your password must be at least ${PASSWORD_MIN_LENGTH} characters long.`;
  if (pw.length > PASSWORD_MAX_LENGTH) return `Your password must be at most ${PASSWORD_MAX_LENGTH} characters long.`;
  const local = email.split('@')[0]?.trim().toLowerCase() ?? '';
  const lower = pw.toLowerCase();
  if (lower === email.trim().toLowerCase() || (local.length >= 3 && lower.includes(local))) {
    return 'Your password must not contain your email address or the part before the @.';
  }
  if (classCount(pw) < 2) {
    return 'Your password must mix at least two kinds of characters: lowercase letters, uppercase letters, digits or symbols.';
  }
  return null;
}
