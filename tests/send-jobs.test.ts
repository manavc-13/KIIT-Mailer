// Pure logic of the send runner (server/lib/jobs.ts), the jobs route limits and the
// public base URL. Nothing here touches the database: jobs.ts builds a Pool at import
// time, which never connects until a query is made, so a dummy DATABASE_URL is enough.
import { describe, it, expect, beforeAll, afterEach, vi } from 'vitest';

process.env.DATABASE_URL ||= 'postgresql://test:test@127.0.0.1:1/never_used';

type Jobs = typeof import('../server/lib/jobs.js');
type Routes = typeof import('../server/routes/jobs.js');
let jobs: Jobs;
let routes: Routes;

beforeAll(async () => {
  jobs = await import('../server/lib/jobs.js');
  routes = await import('../server/routes/jobs.js');
});

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllEnvs();
});

describe('run budget', () => {
  it('defaults to a ~32 s claim window and caps RUN_BUDGET_MS at 35 s', () => {
    expect(jobs.claimBudgetMs(undefined)).toBe(32_000);
    expect(jobs.claimBudgetMs('')).toBe(32_000);
    expect(jobs.claimBudgetMs('abc')).toBe(32_000);
    expect(jobs.claimBudgetMs('-5')).toBe(32_000);
    expect(jobs.claimBudgetMs('10000')).toBe(10_000);
    expect(jobs.claimBudgetMs('40000')).toBe(35_000);
    expect(jobs.claimBudgetMs('999999')).toBe(35_000);
  });

  it('the last claim plus a timed-out send still ends well inside the 60 s function limit', () => {
    const worstSendEnd = jobs.MAX_CLAIM_BUDGET_MS + jobs.SEND_TIMEOUT_MS;
    const recordRetries = jobs.RECORD_SENT_BACKOFF_MS.reduce((a, b) => a + b, 0);
    expect(jobs.FUNCTION_MAX_MS - jobs.MAX_CLAIM_BUDGET_MS).toBeGreaterThanOrEqual(20_000);
    expect(worstSendEnd + recordRetries).toBeLessThan(jobs.FUNCTION_MAX_MS - 5_000);
  });

  it('waits at most 5 s for in-flight messages and never past the claim window', () => {
    expect(jobs.inFlightWaitMs(0, 30_000)).toBe(5_000);
    expect(jobs.inFlightWaitMs(28_000, 30_000)).toBe(2_000);
    expect(jobs.inFlightWaitMs(30_000, 30_000)).toBe(0);
    expect(jobs.inFlightWaitMs(31_000, 30_000)).toBe(0);
  });
});

describe('withTimeout', () => {
  it('passes through a result that arrives in time', async () => {
    await expect(jobs.withTimeout(Promise.resolve('ok'), 1000)).resolves.toBe('ok');
  });

  it('passes through an error that arrives in time', async () => {
    await expect(jobs.withTimeout(Promise.reject(new Error('550 nope')), 1000)).rejects.toThrow('550 nope');
  });

  it('rejects with a SendTimeoutError once the time is up', async () => {
    vi.useFakeTimers();
    const never = new Promise<string>(() => {});
    const p = jobs.withTimeout(never, jobs.SEND_TIMEOUT_MS);
    const check = expect(p).rejects.toBeInstanceOf(jobs.SendTimeoutError);
    await vi.advanceTimersByTimeAsync(jobs.SEND_TIMEOUT_MS);
    await check;
    await expect(p).rejects.toThrow('within 18 seconds');
  });

  it('a late rejection of the abandoned send is not left unhandled', async () => {
    vi.useFakeTimers();
    let rejectLate!: (e: Error) => void;
    const late = new Promise<string>((_, rej) => { rejectLate = rej; });
    const p = jobs.withTimeout(late, 100);
    const check = expect(p).rejects.toBeInstanceOf(jobs.SendTimeoutError);
    await vi.advanceTimersByTimeAsync(100);
    await check;
    rejectLate(new Error('connection closed'));
    await vi.advanceTimersByTimeAsync(0);
  });
});

describe('retryAsync (recording a message Gmail accepted)', () => {
  it('returns at once when the first attempt works', async () => {
    const fn = vi.fn(async () => 'row');
    const wait = vi.fn(async () => {});
    await expect(jobs.retryAsync(fn, [300, 1000], { wait })).resolves.toBe('row');
    expect(fn).toHaveBeenCalledTimes(1);
    expect(wait).not.toHaveBeenCalled();
  });

  it('retries with the given back-off and succeeds on the third attempt', async () => {
    let n = 0;
    const fn = vi.fn(async () => { if (++n < 3) throw new Error('db down'); return 'row'; });
    const waits: number[] = [];
    const errors: number[] = [];
    const r = await jobs.retryAsync(fn, jobs.RECORD_SENT_BACKOFF_MS, {
      wait: async (ms) => { waits.push(ms); },
      onError: (_e, attempt) => errors.push(attempt),
    });
    expect(r).toBe('row');
    expect(fn).toHaveBeenCalledTimes(3);
    expect(waits).toEqual(jobs.RECORD_SENT_BACKOFF_MS);
    expect(errors).toEqual([1, 2]);
  });

  it('gives up after delays.length + 1 attempts with the last error', async () => {
    let n = 0;
    const fn = vi.fn(async () => { throw new Error(`fail ${++n}`); });
    await expect(jobs.retryAsync(fn, [1, 1], { wait: async () => {} })).rejects.toThrow('fail 3');
    expect(fn).toHaveBeenCalledTimes(3);
  });
});

describe('stale claim policy', () => {
  const now = new Date('2026-11-27T10:00:00Z');
  it('a claim older than 3 minutes is stale; a fresh one is not', () => {
    expect(jobs.STALE_CLAIM_MS).toBe(180_000);
    expect(jobs.isStaleClaim(new Date(now.getTime() - 180_001), now)).toBe(true);
    expect(jobs.isStaleClaim(new Date(now.getTime() - 180_000), now)).toBe(false);
    expect(jobs.isStaleClaim(new Date(now.getTime() - 5_000), now)).toBe(false);
  });
  it('a sending row without a claim time is treated as stale', () => {
    expect(jobs.isStaleClaim(null, now)).toBe(true);
  });
  it('stale rows get a plain-English "delivery unknown" error', () => {
    expect(jobs.STALE_SENDING_ERROR).toBe('Delivery status unknown: the server stopped while this message was being sent. Check with the recipient before retrying.');
  });
});

describe('settleAction (what happens to a claimed message that was not sent)', () => {
  const base = { retryable: false, accountLevel: false, attempts: 1, jobStatus: 'sending' };
  it('retryable failures are re-queued while attempts remain, then fail', () => {
    expect(jobs.settleAction({ ...base, retryable: true, attempts: 1 })).toBe('requeue');
    expect(jobs.settleAction({ ...base, retryable: true, attempts: 2 })).toBe('requeue');
    expect(jobs.settleAction({ ...base, retryable: true, attempts: 3 })).toBe('fail');
  });
  it('non-retryable failures fail', () => {
    expect(jobs.settleAction(base)).toBe('fail');
  });
  it('account-level problems (auth, quota, job paused) release the message', () => {
    expect(jobs.settleAction({ ...base, accountLevel: true })).toBe('release');
    expect(jobs.settleAction({ ...base, accountLevel: true, jobStatus: 'paused' })).toBe('release');
  });
  it('cancelled while in flight: retryable or account-level failures are skipped, never re-queued', () => {
    expect(jobs.settleAction({ ...base, retryable: true, jobStatus: 'cancelled' })).toBe('skip');
    expect(jobs.settleAction({ ...base, retryable: true, attempts: 3, jobStatus: 'cancelled' })).toBe('skip');
    expect(jobs.settleAction({ ...base, accountLevel: true, jobStatus: 'cancelled' })).toBe('skip');
    expect(jobs.CANCELLED_ERROR).toBe('Cancelled before sending.');
  });
  it('cancelled while in flight: a final (non-retryable) failure still counts as failed', () => {
    expect(jobs.settleAction({ ...base, jobStatus: 'cancelled' })).toBe('fail');
  });
});

describe('APP_BASE_URL resolution', () => {
  it('prefers APP_BASE_URL and strips trailing slashes', async () => {
    const { appBaseUrl } = await import('../server/lib/unsubscribe.js');
    vi.stubEnv('APP_BASE_URL', 'https://mailer.example///');
    vi.stubEnv('VERCEL_PROJECT_PRODUCTION_URL', 'kiit-mailer.vercel.app');
    expect(appBaseUrl()).toBe('https://mailer.example');
  });
  it('falls back to the Vercel production domain', async () => {
    const { appBaseUrl } = await import('../server/lib/unsubscribe.js');
    vi.stubEnv('APP_BASE_URL', '');
    vi.stubEnv('VERCEL_PROJECT_PRODUCTION_URL', 'kiit-mailer.vercel.app');
    expect(appBaseUrl()).toBe('https://kiit-mailer.vercel.app');
    vi.stubEnv('VERCEL_PROJECT_PRODUCTION_URL', 'https://kiit-mailer.vercel.app/');
    expect(appBaseUrl()).toBe('https://kiit-mailer.vercel.app');
  });
  it('falls back to the local dev server', async () => {
    const { appBaseUrl } = await import('../server/lib/unsubscribe.js');
    vi.stubEnv('APP_BASE_URL', '');
    vi.stubEnv('VERCEL_PROJECT_PRODUCTION_URL', '');
    expect(appBaseUrl()).toBe('http://localhost:5173');
  });
  it('hosted-image detection follows the same base URL', async () => {
    const { hostedImageIds } = await import('../server/lib/send/images.js');
    const id = '0b8f3c1e-1111-4222-8333-944455556666';
    vi.stubEnv('APP_BASE_URL', '');
    vi.stubEnv('VERCEL_PROJECT_PRODUCTION_URL', 'kiit-mailer.vercel.app');
    expect(hostedImageIds(`<img src="https://kiit-mailer.vercel.app/api/uploads/${id}/content">`)).toEqual([id]);
    expect(hostedImageIds(`<img src="/api/uploads/${id}/content">`)).toEqual([id]);
    expect(hostedImageIds(`<img src="https://elsewhere.example/api/uploads/${id}/content">`)).toEqual([]);
  });
});

describe('error codes', () => {
  it('quota and auth errors carry stable codes', async () => {
    const { quotaExhausted, authFailed, ApiError } = await import('../server/lib/http.js');
    const q = quotaExhausted('No quota left.', { remaining: 0 });
    expect(q).toBeInstanceOf(ApiError);
    expect(q.status).toBe(400);
    expect(q.code).toBe('quota_exhausted');
    expect(q.message).toBe('No quota left.');
    expect(authFailed('App Password rejected.').code).toBe('auth_failed');
  });
});

describe('jobs route request limits', () => {
  const sender = '6a1b2c3d-1111-4222-8333-944455556666';
  const valid = () => ({
    subject: 'Hello', editorMode: 'rich', bodyHtml: '<p>Hi</p>', theme: 'plain',
    senderAccountId: sender, mode: 'personal', recipients: [{ to: 'a@kiit.ac.in', data: { name: 'A' } }],
  });
  const firstMessage = (r: { success: boolean; error?: { issues: { message: string }[] } }) => r.error?.issues[0]?.message;

  it('accepts a normal request', () => {
    expect(routes.createJobSchema.safeParse(valid()).success).toBe(true);
  });

  it('caps the HTML body at 1 MB (counted in UTF-8 bytes)', () => {
    const ok = routes.createJobSchema.safeParse({ ...valid(), bodyHtml: 'a'.repeat(routes.MAX_BODY_HTML_BYTES) });
    expect(ok.success).toBe(true);
    const big = routes.createJobSchema.safeParse({ ...valid(), bodyHtml: 'a'.repeat(routes.MAX_BODY_HTML_BYTES + 1) });
    expect(big.success).toBe(false);
    expect(firstMessage(big)).toMatch(/email body is too large \(over 1 MB\)/);
    // 400k three-byte characters = 1.2 MB even though it is under a million characters.
    expect(routes.createJobSchema.safeParse({ ...valid(), bodyHtml: 'ଓ'.repeat(400_000) }).success).toBe(false);
  });

  it('caps the plain-text body at 200 KB', () => {
    expect(routes.createJobSchema.safeParse({ ...valid(), bodyText: 'a'.repeat(200 * 1024) }).success).toBe(true);
    const big = routes.createJobSchema.safeParse({ ...valid(), bodyText: 'a'.repeat(200 * 1024 + 1) });
    expect(big.success).toBe(false);
    expect(firstMessage(big)).toMatch(/plain-text version .* too large/);
  });

  it('caps recipients at 5,000 rows', () => {
    const rows = (n: number) => Array.from({ length: n }, (_, i) => ({ to: `p${i}@kiit.ac.in` }));
    expect(routes.MAX_RECIPIENT_ROWS).toBe(5000);
    expect(routes.createJobSchema.safeParse({ ...valid(), recipients: rows(5000) }).success).toBe(true);
    const big = routes.createJobSchema.safeParse({ ...valid(), recipients: rows(5001) });
    expect(big.success).toBe(false);
    expect(firstMessage(big)).toMatch(/Too many recipients .*5,000/);
  });

  it('caps per-field recipient data', () => {
    const r = (rec: object) => routes.createJobSchema.safeParse({ ...valid(), recipients: [{ to: 'a@kiit.ac.in', ...rec }] });
    expect(firstMessage(r({ to: 'a'.repeat(2001) }))).toMatch(/To field is too long/);
    expect(firstMessage(r({ cc: 'a'.repeat(5001) }))).toMatch(/CC field is too long/);
    expect(firstMessage(r({ bcc: 'a'.repeat(10_001) }))).toMatch(/BCC field is too long/);
    expect(firstMessage(r({ data: { name: 'x'.repeat(5001) } }))).toMatch(/value in the recipient list is too long/);
    expect(firstMessage(r({ data: { ['k'.repeat(101)]: 'x' } }))).toMatch(/column name .* too long/);
    const wide = Object.fromEntries(Array.from({ length: 101 }, (_, i) => [`c${i}`, 'x']));
    expect(firstMessage(r({ data: wide }))).toMatch(/too many columns/);
  });

  it('caps the whole recipient list so the request stays under Vercel\'s 4.5 MB body limit', () => {
    const fat = Array.from({ length: 5000 }, (_, i) => ({ to: `p${i}@kiit.ac.in`, data: { a: 'x'.repeat(600) } }));
    const big = routes.createJobSchema.safeParse({ ...valid(), recipients: fat });
    expect(big.success).toBe(false);
    expect(firstMessage(big)).toMatch(/recipient list is too large/);
    const worst = routes.MAX_BODY_HTML_BYTES * 1.25 + routes.MAX_BODY_TEXT_BYTES * 1.25 + routes.MAX_RECIPIENT_LIST_BYTES + 200 * 1024;
    expect(worst).toBeLessThan(4.5 * 1024 * 1024);
  });

  it('test sends use the same compose limits', () => {
    const t = { subject: 'S', editorMode: 'rich', bodyHtml: 'a'.repeat(routes.MAX_BODY_HTML_BYTES + 1), theme: 'plain', senderAccountId: sender, to: ['a@kiit.ac.in'] };
    expect(routes.testSendSchema.safeParse(t).success).toBe(false);
    expect(routes.testSendSchema.safeParse({ ...t, bodyHtml: '<p>x</p>' }).success).toBe(true);
  });
});
