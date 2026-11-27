import { describe, it, expect, afterEach } from 'vitest';
import {
  mergeCopies, planPersonal, planBroadcast, totalRecipients, resolveUnsubscribe, checkQuota, externalRecipientCount, paceDelayMs, firstAddress,
} from '../server/lib/send/plan.js';
import { parseEmailList } from '../shared/text.js';

describe('mergeCopies (legacy CC/BCC column rules)', () => {
  it('unions row and global CC, dedupes case-insensitively and never repeats To', () => {
    const r = mergeCopies(['a@x.com'], 'b@x.com; A@x.com', null, 'B@X.com, c@x.com', null);
    expect(r.cc).toEqual(['b@x.com', 'c@x.com']);
  });
  it('BCC excludes To and final CC', () => {
    const r = mergeCopies(['a@x.com'], 'b@x.com', 'a@x.com, b@x.com, d@x.com', null, 'e@x.com');
    expect(r.cc).toEqual(['b@x.com']);
    expect(r.bcc).toEqual(['d@x.com', 'e@x.com']);
  });
  it('resolves placeholders in global CC/BCC against the row', () => {
    const r = mergeCopies(['a@x.com'], '', '', '{Supervisor}', '', { Supervisor: 'boss@x.com' });
    expect(r.cc).toEqual(['boss@x.com']);
  });
});

describe('planPersonal', () => {
  it('creates one message per entry with merged copies and counts recipients', () => {
    const p = planPersonal([
      { to: 'a@x.com', cc: 'c1@x.com', data: { Name: 'A' } },
      { to: 'b@x.com' },
    ], { globalCc: 'g@x.com', globalBcc: 'h@x.com' });
    expect(p.messages).toHaveLength(2);
    expect(p.messages[0]).toMatchObject({ to: 'a@x.com', cc: 'c1@x.com, g@x.com', bcc: 'h@x.com', recipientCount: 4, rowData: { Name: 'A' } });
    expect(p.messages[1].recipientCount).toBe(3);
    expect(totalRecipients(p.messages)).toBe(7);
  });
  it('reports invalid addresses and empty rows', () => {
    const p = planPersonal([{ to: 'bad-address' }, { to: '' }, { to: 'ok@x.com', cc: 'nope@' }]);
    expect(p.invalid).toEqual(['bad-address', 'nope@']);
    expect(p.emptyRows).toEqual([2]);
  });
  it('dedupes To addresses only when asked', () => {
    const rs = [{ to: 'a@x.com' }, { to: 'A@x.com' }, { to: 'b@x.com' }];
    expect(planPersonal(rs).messages).toHaveLength(3);
    const d = planPersonal(rs, { dedupe: true });
    expect(d.messages).toHaveLength(2);
    expect(d.duplicatesRemoved).toBe(1);
  });
  it('suppression: a suppressed To becomes a skipped message; suppressed copies are dropped', () => {
    const suppressed = new Set(['gone@x.com', 'cc-gone@x.com']);
    const p = planPersonal([
      { to: 'GONE@x.com', cc: 'keep@x.com' },
      { to: 'a@x.com', cc: 'cc-gone@x.com, keep@x.com' },
    ], { suppressed });
    expect(p.messages[0].skipReason).toBeTruthy();
    expect(p.messages[0].recipientCount).toBe(0);
    expect(p.messages[1]).toMatchObject({ to: 'a@x.com', cc: 'keep@x.com', recipientCount: 2 });
    expect(p.suppressedAddresses.sort()).toEqual(['cc-gone@x.com', 'gone@x.com']);
    expect(totalRecipients(p.messages)).toBe(2);
  });
});

describe('planBroadcast (BCC chunking)', () => {
  const many = (n: number, domain = 'ext.com') => Array.from({ length: n }, (_, i) => ({ to: `user${i}@${domain}` }));

  it('chunks into messages of at most 90 BCC with To = the sender', () => {
    const p = planBroadcast(many(250), 'iqac@kiit.ac.in');
    expect(p.messages).toHaveLength(3);
    for (const m of p.messages) {
      expect(m.to).toBe('iqac@kiit.ac.in');
      expect(parseEmailList(m.bcc).length).toBeLessThanOrEqual(90);
      expect(m.rowData).toEqual({});
    }
    const all = p.messages.flatMap(m => parseEmailList(m.bcc));
    expect(new Set(all).size).toBe(250);
    expect(totalRecipients(p.messages)).toBe(250 + 3);
  });
  it('dedupes, includes global BCC and per-entry cc/bcc, drops the sender itself', () => {
    const p = planBroadcast([{ to: 'a@x.com', cc: 'b@x.com' }, { to: 'A@x.com', bcc: 'iqac@kiit.ac.in' }], 'iqac@kiit.ac.in', { globalBcc: 'c@x.com, b@x.com' });
    expect(p.messages).toHaveLength(1);
    expect(parseEmailList(p.messages[0].bcc).sort()).toEqual(['a@x.com', 'b@x.com', 'c@x.com']);
  });
  it('global CC goes on the first message only and shares its 90 budget', () => {
    const p = planBroadcast(many(200), 'iqac@kiit.ac.in', { globalCc: 'dean@kiit.ac.in, vc@kiit.ac.in' });
    expect(p.messages[0].cc).toBe('dean@kiit.ac.in, vc@kiit.ac.in');
    expect(parseEmailList(p.messages[0].bcc).length + 2).toBeLessThanOrEqual(90);
    expect(p.messages.slice(1).every(m => m.cc === null)).toBe(true);
    expect(p.messages.flatMap(m => parseEmailList(m.bcc))).toHaveLength(200);
  });
  it('removes suppressed addresses', () => {
    const p = planBroadcast(many(5), 'iqac@kiit.ac.in', { suppressed: new Set(['user1@ext.com']) });
    expect(parseEmailList(p.messages[0].bcc)).not.toContain('user1@ext.com');
    expect(p.suppressedAddresses).toEqual(['user1@ext.com']);
  });
});

describe('quota math', () => {
  it('passes when the send fits', () => {
    expect(checkQuota(800, 2000, 1200, 'x@kiit.ac.in')).toEqual({ ok: true, remaining: 800, message: null });
  });
  it('fails with numbers when it does not', () => {
    const q = checkQuota(2340, 2000, 800, 'x@kiit.ac.in');
    expect(q.ok).toBe(false);
    expect(q.remaining).toBe(1200);
    expect(q.message).toBe('This send needs 2,340 recipients but only 1,200 remain today for x@kiit.ac.in. Split the list or use another account.');
  });
  it('never reports negative remaining', () => {
    expect(checkQuota(1, 100, 150, 'x@kiit.ac.in').remaining).toBe(0);
  });
});

describe('unsubscribe auto rule', () => {
  const msgs = (n: number, domain: string) => planPersonal(Array.from({ length: n }, (_, i) => ({ to: `p${i}@${domain}` }))).messages;
  it('is on with more than 50 external recipients', () => {
    expect(resolveUnsubscribe('auto', msgs(51, 'gmail.com'))).toBe(true);
    expect(resolveUnsubscribe('auto', msgs(50, 'gmail.com'))).toBe(false);
  });
  it('ignores kiit.ac.in recipients', () => {
    expect(externalRecipientCount(msgs(200, 'kiit.ac.in'))).toBe(0);
    expect(resolveUnsubscribe('auto', msgs(200, 'kiit.ac.in'))).toBe(false);
  });
  it('respects explicit settings', () => {
    expect(resolveUnsubscribe(true, [])).toBe(true);
    expect(resolveUnsubscribe(false, msgs(80, 'gmail.com'))).toBe(false);
  });
});

describe('helpers', () => {
  afterEach(() => { delete process.env.SEND_DELAY_MIN_MS; delete process.env.SEND_DELAY_MAX_MS; });
  it('pace delay defaults to 1500-3000 ms and honours env overrides', () => {
    expect(paceDelayMs(() => 0)).toBe(1500);
    expect(paceDelayMs(() => 1)).toBe(3000);
    process.env.SEND_DELAY_MIN_MS = '0'; process.env.SEND_DELAY_MAX_MS = '10';
    expect(paceDelayMs(() => 0.5)).toBe(5);
  });
  it('firstAddress', () => {
    expect(firstAddress('a@x.com, b@x.com')).toBe('a@x.com');
    expect(firstAddress('')).toBeNull();
  });
});
