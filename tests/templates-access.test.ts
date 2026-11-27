import { describe, expect, it } from 'vitest';
import {
  copyName, MAX_HTML_BYTES, parseImportItem, resolveImportNames, saveTemplateSchema, shareSchema, templateAccess,
} from '../server/lib/templates';
import type { TemplateVisibility } from '../shared/types';

const owner = { id: 'owner', role: 'sender' as const };
const other = { id: 'other', role: 'sender' as const };
const admin = { id: 'admin', role: 'admin' as const };
const tpl = (visibility: TemplateVisibility) => ({ ownerId: 'owner', visibility });

describe('templateAccess (visibility rule)', () => {
  it('owner can do everything except share (unless admin)', () => {
    for (const v of ['private', 'shared_all', 'shared_users'] as const) {
      expect(templateAccess(tpl(v), owner, false)).toEqual({
        listed: true, view: true, edit: true, delete: true, share: false, seeShares: true,
      });
    }
  });

  it('other users: private hidden, shared_all listed, shared_users only when in the share list', () => {
    expect(templateAccess(tpl('private'), other, false).view).toBe(false);
    expect(templateAccess(tpl('private'), other, true).view).toBe(false); // stale share row does not leak a private template
    expect(templateAccess(tpl('shared_all'), other, false)).toMatchObject({ listed: true, view: true, edit: false, delete: false, share: false, seeShares: false });
    expect(templateAccess(tpl('shared_users'), other, false)).toMatchObject({ listed: false, view: false });
    expect(templateAccess(tpl('shared_users'), other, true)).toMatchObject({ listed: true, view: true, edit: false, delete: false });
  });

  it("admins list like everyone else, but may open, delete and share any template; never edit others'", () => {
    const a = templateAccess(tpl('private'), admin, false);
    expect(a).toEqual({ listed: false, view: true, edit: false, delete: true, share: true, seeShares: true });
    expect(templateAccess(tpl('shared_all'), admin, false).listed).toBe(true);
    expect(templateAccess({ ownerId: 'admin', visibility: 'private' }, admin, false))
      .toEqual({ listed: true, view: true, edit: true, delete: true, share: true, seeShares: true });
  });
});

describe('import naming', () => {
  it('suffixes names that clash with the owner\'s templates or earlier imports (case-insensitive)', () => {
    expect(resolveImportNames(['Circular', 'notice'], ['Circular', 'Notice', 'New', 'new', 'Circular']))
      .toEqual(['Circular (imported)', 'Notice (imported)', 'New', 'new (imported)', 'Circular (imported 2)']);
  });
  it('keeps names within 120 characters', () => {
    const long = 'x'.repeat(120);
    const [n] = resolveImportNames([long], [long]);
    expect(n.length).toBeLessThanOrEqual(120);
    expect(n.endsWith(' (imported)')).toBe(true);
    expect(copyName(long)).toHaveLength(120);
    expect(copyName('Invite')).toBe('Invite (copy)');
  });
});

describe('save validation', () => {
  const ok = { name: ' Invite ', subject: 'Hi', editorMode: 'rich', html: '<p>x</p>', theme: 'heritage' };
  it('accepts a minimal template and fills defaults', () => {
    const r = saveTemplateSchema.parse(ok);
    expect(r).toMatchObject({ name: 'Invite', category: null, preheader: '', docJson: null, text: null, includeSignature: true, defaultSenderAccountId: null });
  });
  it('rejects bad names, themes, oversize html and bad sender ids with readable messages', () => {
    const msg = (v: unknown) => { const r = saveTemplateSchema.safeParse(v); return r.success ? null : r.error.issues[0].message; };
    expect(msg({ ...ok, name: '   ' })).toMatch(/name/i);
    expect(msg({ ...ok, name: 'x'.repeat(121) })).toMatch(/120/);
    expect(msg({ ...ok, theme: 'neon' })).toMatch(/Theme must be one of/);
    expect(msg({ ...ok, html: 'x'.repeat(MAX_HTML_BYTES + 1) })).toMatch(/500 KB/);
    expect(msg({ ...ok, html: 'x'.repeat(MAX_HTML_BYTES) })).toBeNull();
    expect(msg({ ...ok, defaultSenderAccountId: 'nope' })).toMatch(/valid id/);
  });
  it('names the failing template in import errors', () => {
    const r = parseImportItem({ ...ok, name: 'Bad', theme: 'neon' }, 2);
    expect(r).toEqual({ ok: false, error: expect.stringMatching(/^Template 3 \("Bad"\): Theme/) });
  });
  it('share requires people for shared_users', () => {
    expect(shareSchema.safeParse({ visibility: 'shared_users', userIds: [] }).success).toBe(false);
    expect(shareSchema.safeParse({ visibility: 'shared_all' }).success).toBe(true);
  });
});
