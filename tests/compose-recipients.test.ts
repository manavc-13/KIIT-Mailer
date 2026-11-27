import { describe, it, expect } from 'vitest';
import {
  detectRoleMap, canonicalizeRow, resolveCopies, computeValidation, buildJobRecipients, dedupeRows,
  placeholderHeaders, bccMessageCount, chunk, sanitizeRoleMap, SINGLE_ROLE_MAP, type RoleMap,
} from '../shared/recipients';

const RM: RoleMap = { email: 'Email', name: 'Name', cc: 'CC', bcc: 'BCC' };

describe('detectRoleMap', () => {
  it('detects aliases case-, whitespace- and BOM-insensitively', () => {
    const m = detectRoleMap(['﻿E-mail Address', ' Full Name ', 'Copy', 'BLIND COPY', 'Dept']);
    expect(m).toEqual({ email: '﻿E-mail Address', name: ' Full Name ', cc: 'Copy', bcc: 'BLIND COPY' });
  });
  it('returns null for roles with no matching column', () => {
    expect(detectRoleMap(['Dept', 'Phone'])).toEqual({ email: null, name: null, cc: null, bcc: null });
  });
  it('prefers a real Email column over To', () => {
    expect(detectRoleMap(['To', 'Email']).email).toBe('Email');
    expect(detectRoleMap(['To', 'Name']).email).toBe('To');
  });
  it('sanitizeRoleMap drops stale columns', () => {
    expect(sanitizeRoleMap(RM, ['Name', 'Email'])).toEqual({ email: 'Email', name: 'Name', cc: null, bcc: null });
  });
});

describe('canonicalizeRow', () => {
  it('adds Email/Name aliases from mapped columns', () => {
    const row = { 'E-mail': ' a@x.com ', 'Full Name': ' Asha ', Dept: 'CSE' };
    const c = canonicalizeRow(row, { email: 'E-mail', name: 'Full Name', cc: null, bcc: null });
    expect(c).toMatchObject({ Email: 'a@x.com', Name: 'Asha', Dept: 'CSE', 'E-mail': ' a@x.com ' });
  });
});

describe('resolveCopies', () => {
  it('unions row and global copies, de-dupes case-insensitively and never repeats To', () => {
    const row = { Email: 'to@x.com', CC: 'hod@x.com; TO@x.com', BCC: 'audit@x.com' };
    const r = resolveCopies(row, RM, 'HOD@x.com, dean@x.com', 'audit@x.com, dean@x.com, to@x.com', 'to@x.com');
    expect(r.cc).toBe('hod@x.com, dean@x.com');
    // BCC excludes To and anything already in CC
    expect(r.bcc).toBe('audit@x.com');
    expect(r.invalid).toEqual([]);
  });
  it('resolves placeholders in global fields and reports invalid addresses', () => {
    const row = { Email: 'a@x.com', Supervisor: 'sup@x.com', CC: 'not-an-email' };
    const r = resolveCopies(row, RM, '{Supervisor}', '{Missing}', 'a@x.com');
    expect(r.cc).toBe('sup@x.com');
    expect(r.bcc).toBe('');
    expect(r.invalid).toEqual(expect.arrayContaining(['not-an-email', '{Missing}']));
  });
  it('ignores row copies when the role is unmapped', () => {
    const r = resolveCopies({ Email: 'a@x.com', CC: 'c@x.com' }, SINGLE_ROLE_MAP, '', '', 'a@x.com');
    expect(r.cc).toBe('');
  });
});

describe('computeValidation', () => {
  it('counts valid, invalid, duplicates, missing email and missing name', () => {
    const rows = [
      { Name: 'A', Email: 'a@x.com' },
      { Name: '', Email: 'A@x.com' },
      { Name: 'B', Email: 'bad' },
      { Name: 'C', Email: '' },
      { Name: '', Email: '' }, // blank row ignored
    ];
    const v = computeValidation(rows, RM);
    expect(v.total).toBe(4);
    expect(v.valid).toBe(2);
    expect(v.invalid).toEqual(['bad']);
    expect(v.duplicates).toEqual(['a@x.com']);
    expect(v.duplicateRows).toBe(1);
    expect(v.missingEmail).toBe(1);
    expect(v.missingName).toBe(1);
    expect(v.uniqueCount).toBe(1);
  });
});

describe('dedupeRows', () => {
  it('keeps the first row per address', () => {
    const d = dedupeRows([{ Email: 'a@x.com', N: '1' }, { Email: 'A@X.com', N: '2' }, { Email: 'b@x.com', N: '3' }], RM);
    expect(d.rows.map(r => r.N)).toEqual(['1', '3']);
    expect(d.dropped).toBe(1);
  });
});

describe('placeholderHeaders', () => {
  it('lists canonical Name/Email first, then columns, without duplicates', () => {
    expect(placeholderHeaders(['E-mail', 'Full Name', 'Dept', 'name'], { email: 'E-mail', name: 'Full Name', cc: null, bcc: null }))
      .toEqual(['Name', 'Email', 'E-mail', 'Full Name', 'Dept']);
  });
});

describe('buildJobRecipients: personal', () => {
  const rows = [
    { Name: 'A', Email: 'a@x.com', CC: 'c1@x.com', BCC: '' },
    { Name: 'A again', Email: 'A@x.com', CC: '', BCC: '' },
    { Name: 'B', Email: 'b@x.com', CC: '', BCC: 'b-bcc@x.com' },
    { Name: 'Bad', Email: 'nope', CC: '', BCC: '' },
    { Name: 'NoMail', Email: '', CC: '', BCC: '' },
  ];
  it('builds one entry per valid row with resolved copies and canonical data', () => {
    const r = buildJobRecipients({ rows, roleMap: RM, mode: 'personal', globalCc: 'g@x.com', globalBcc: '' });
    expect(r.recipients).toHaveLength(3);
    expect(r.recipients[0]).toEqual({ to: 'a@x.com', cc: 'c1@x.com, g@x.com', data: expect.objectContaining({ Name: 'A', Email: 'a@x.com' }) });
    expect(r.recipients[2]).toMatchObject({ to: 'b@x.com', cc: 'g@x.com', bcc: 'b-bcc@x.com' });
    expect(r.skippedInvalid).toEqual(['nope']);
    expect(r.skippedMissing).toBe(1);
    expect(r.messageCount).toBe(3);
    expect(r.ccTotal).toBe(4);
    expect(r.bccTotal).toBe(1);
    expect(r.quotaCost).toBe(3 + 4 + 1);
  });
  it('dedupes when asked', () => {
    const r = buildJobRecipients({ rows, roleMap: RM, mode: 'personal', dedupe: true });
    expect(r.recipients.map(x => x.to)).toEqual(['a@x.com', 'b@x.com']);
    expect(r.droppedDuplicates).toBe(1);
  });
});

describe('buildJobRecipients: bcc_broadcast', () => {
  it('produces a flat de-duplicated address list and chunked message count', () => {
    const rows = Array.from({ length: 200 }, (_, i) => ({ Email: `u${i % 185}@x.com`, CC: i === 0 ? 'row-cc@x.com' : '' }));
    const r = buildJobRecipients({ rows, roleMap: RM, mode: 'bcc_broadcast', globalCc: 'cc@x.com, u1@x.com', globalBcc: 'bcc@x.com' });
    expect(r.recipients).toHaveLength(185);
    expect(r.recipients[0]).toEqual({ to: 'u0@x.com' });
    expect(r.recipients.every(x => !x.data && !x.cc)).toBe(true);
    expect(r.droppedDuplicates).toBe(15);
    expect(r.messageCount).toBe(3); // 90 + 90 + 5
    expect(r.ignoredRowCopies).toBe(true);
    // global CC minus list members applies to every chunk
    expect(r.ccTotal).toBe(3);
    expect(r.bccTotal).toBe(185 + 3);
    expect(r.quotaCost).toBe(3 + 3 + 188);
  });
  it('bccMessageCount and chunk respect the 90-per-message cap', () => {
    expect(bccMessageCount(0)).toBe(0);
    expect(bccMessageCount(90)).toBe(1);
    expect(bccMessageCount(91)).toBe(2);
    expect(chunk([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]]);
  });
});
