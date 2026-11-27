import { describe, expect, it } from 'vitest';
import { hashPassword, passwordPolicyChecks, validatePasswordPolicy, verifyPassword } from '../server/lib/passwords.js';

process.env.DATABASE_URL ||= 'postgres://test:test@localhost:5432/test';

describe('validatePasswordPolicy', () => {
  const email = 'iqac.office@kiit.ac.in';

  it('accepts a reasonable password', () => {
    expect(validatePasswordPolicy('Monsoon-Library-42', email)).toBeNull();
    expect(validatePasswordPolicy('correct horse battery 9', email)).toBeNull();
  });

  it('requires at least 10 characters', () => {
    expect(validatePasswordPolicy('Ab1!xyz', email)).toMatch(/at least 10/);
    expect(validatePasswordPolicy('Ab1!xyzabc', email)).toBeNull();
  });

  it('rejects passwords containing the email local part or equal to the email', () => {
    expect(validatePasswordPolicy('IQAC.OFFICE2025', email)).toMatch(/email/);
    expect(validatePasswordPolicy('my-iqac.office-pw', email)).toMatch(/email/);
    expect(validatePasswordPolicy(email, email)).toMatch(/email/);
  });

  it('rejects a single character class', () => {
    expect(validatePasswordPolicy('abcdefghijkl', email)).toMatch(/two kinds/);
    expect(validatePasswordPolicy('123456789012', email)).toMatch(/two kinds/);
    expect(validatePasswordPolicy('ABCDEFGHIJKL', email)).toMatch(/two kinds/);
    expect(validatePasswordPolicy('!@#$%^&*()_+', email)).toMatch(/two kinds/);
  });

  it('exposes per-rule checks for live hints', () => {
    const checks = passwordPolicyChecks('short', email);
    expect(checks.map(c => c.id)).toEqual(['length', 'email', 'variety']);
    expect(checks.find(c => c.id === 'length')?.ok).toBe(false);
    expect(passwordPolicyChecks('Monsoon-Library-42', email).every(c => c.ok)).toBe(true);
  });
});

describe('scrypt hashing', () => {
  it('hashes with scrypt and verifies', async () => {
    const h = await hashPassword('Monsoon-Library-42');
    expect(h.startsWith('$scrypt$ln=15,r=8,p=1$')).toBe(true);
    expect(await verifyPassword(h, 'Monsoon-Library-42')).toBe(true);
    expect(await verifyPassword(h, 'monsoon-library-42')).toBe(false);
  });

  it('salts every hash', async () => {
    const a = await hashPassword('Monsoon-Library-42');
    const b = await hashPassword('Monsoon-Library-42');
    expect(a).not.toBe(b);
    expect(await verifyPassword(b, 'Monsoon-Library-42')).toBe(true);
  });

  it('rejects a tampered hash', async () => {
    const h = await hashPassword('Monsoon-Library-42');
    const parts = h.split('$');
    const key = Buffer.from(parts[4], 'base64'); key[0] ^= 1;
    parts[4] = key.toString('base64');
    expect(await verifyPassword(parts.join('$'), 'Monsoon-Library-42')).toBe(false);
  });

  it('returns false for a malformed hash instead of throwing', async () => {
    expect(await verifyPassword('not-a-hash', 'whatever')).toBe(false);
  });
});

describe('app password helpers', () => {
  it('normalizes the four-groups-of-four format and rejects other shapes', async () => {
    const { normalizeAppPassword } = await import('../server/lib/senders.js');
    expect(normalizeAppPassword('abcd efgh ijkl mnop')).toEqual({ ok: true, value: 'abcdefghijklmnop' });
    expect(normalizeAppPassword(' abcdefghijklmnop ')).toEqual({ ok: true, value: 'abcdefghijklmnop' });
    expect(normalizeAppPassword('abcd efgh ijkl').ok).toBe(false);
    expect(normalizeAppPassword('abcd-efgh-ijkl-mnop').ok).toBe(false);
    expect(normalizeAppPassword('abcd efgh ijkl mn0p').ok).toBe(false);
  });

  it('explains SMTP errors in plain English', async () => {
    const { explainSmtpError } = await import('../server/lib/senders.js');
    const e = 'office@kiit.ac.in';
    expect(explainSmtpError({ responseCode: 535, response: '535-5.7.8 Username and Password not accepted. BadCredentials' }, e)).toMatch(/rejected the App Password/);
    expect(explainSmtpError({ responseCode: 534, response: '534-5.7.9 Application-specific password required.' }, e)).toMatch(/2-Step Verification/);
    expect(explainSmtpError({ code: 'ETIMEDOUT', message: 'Connection timeout' }, e)).toMatch(/Could not reach smtp\.gmail\.com/);
    expect(explainSmtpError({ code: 'VERIFY_TIMEOUT' }, e)).toMatch(/15 seconds/);
  });
});
