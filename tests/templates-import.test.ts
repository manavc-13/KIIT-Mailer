import { describe, expect, it } from 'vitest';
import {
  buildExport, legacyToRequest, legacyToRequests, parseImport, parseImportText, templateToRequest, chunk, copyNameFor,
} from '../src/templates/importExport';
import type { Template } from '../shared/types';

describe('legacy (v1) import mapping', () => {
  it("maps editorMode 'html' to html mode with the raw html", () => {
    const r = legacyToRequest('Circular', { subject: 'S', preheader: 'P', editorMode: 'html', html: '<table>x</table>', text: 'ignored', richHtml: '<p>no</p>' });
    expect(r).toMatchObject({ name: 'Circular', subject: 'S', preheader: 'P', editorMode: 'html', html: '<table>x</table>', theme: 'plain', docJson: null });
  });

  it("maps editorMode 'rich' to rich mode from richHtml with docJson null", () => {
    const r = legacyToRequest('Rich', { subject: 'S', editorMode: 'rich', html: '<b>old html</b>', richHtml: '<p><strong>Hi</strong></p>' });
    expect(r.editorMode).toBe('rich');
    expect(r.docJson).toBeNull();
    expect(r.html).toBe('<p><strong>Hi</strong></p>');
    expect(r.theme).toBe('plain');
  });

  it("maps editorMode 'text' to rich mode with escaped text wrapped in <p>, keeping the text part", () => {
    const text = 'Dear {Name},\n\nMeeting at 5 <sharp> & on time.';
    const r = legacyToRequest('Plain', { subject: 'S', editorMode: 'text', text });
    expect(r.editorMode).toBe('rich');
    expect(r.text).toBe(text);
    expect(r.html).toBe('<p>Dear {Name},<br>&nbsp;<br>Meeting at 5 &lt;sharp&gt; &amp; on time.</p>');
    expect(r.theme).toBe('plain');
  });

  it('defaults a missing editorMode to html (v1 default) and tolerates missing fields', () => {
    const r = legacyToRequest('  Bare  ', {});
    expect(r).toMatchObject({ name: 'Bare', subject: '', preheader: '', editorMode: 'html', html: '' });
  });

  it('maps a whole v1 object and parseImport detects it', () => {
    const v1 = {
      A: { subject: 'a', editorMode: 'html', html: '<p>a</p>' },
      B: { subject: 'b', editorMode: 'text', text: 'b' },
    };
    expect(legacyToRequests(v1).map(t => t.name)).toEqual(['A', 'B']);
    const parsed = parseImport(v1);
    expect(parsed.format).toBe('legacy');
    expect(parsed.templates).toHaveLength(2);
  });
});

const sample: Template = {
  id: '00000000-0000-4000-8000-000000000001', name: 'Invitation', category: 'Events', subject: 'You are invited',
  theme: 'heritage', editorMode: 'rich', visibility: 'shared_all', ownerId: 'u1', ownerName: 'Owner', isOwner: true,
  defaultSenderAccountId: null, updatedAt: '2026-11-01T00:00:00.000Z', preheader: 'Join us',
  docJson: { type: 'doc', content: [{ type: 'paragraph' }] }, html: '<p>Hello {Name}</p>', text: null,
  includeSignature: true, sharedUserIds: ['u2'],
};

describe('v2 export / import', () => {
  it('round-trips through JSON', () => {
    const file = buildExport([sample, { ...sample, id: 'x', name: 'Raw', editorMode: 'html', html: '<html></html>', theme: 'plain', category: null, text: 'txt' }], new Date('2026-11-27T00:00:00Z'));
    expect(file.version).toBe(2);
    expect(file.exportedAt).toBe('2026-11-27T00:00:00.000Z');
    const back = parseImportText(JSON.stringify(file));
    expect(back.format).toBe('v2');
    expect(back.templates).toEqual(file.templates);
    expect(back.templates[0]).toEqual(templateToRequest(sample));
    // Server-only fields never leak into the export.
    expect(Object.keys(back.templates[0])).not.toContain('id');
    expect(Object.keys(back.templates[0])).not.toContain('sharedUserIds');
  });

  it('normalises unknown themes and modes in v2 files', () => {
    const r = parseImport({ version: 2, templates: [{ name: 'X', html: '', editorMode: 'weird', theme: 'neon' }] });
    expect(r.templates[0]).toMatchObject({ editorMode: 'rich', theme: 'plain', includeSignature: true, category: null });
  });

  it('rejects files that are neither format', () => {
    expect(() => parseImportText('not json')).toThrow(/valid JSON/);
    expect(() => parseImport({ foo: 1 })).toThrow(/not a KIIT Mailer template export/);
    expect(() => parseImport({ version: 2, templates: [{ html: '' }] })).toThrow(/no name/);
  });

  it('chunks imports to the API limit and names copies within 120 chars', () => {
    expect(chunk(Array.from({ length: 250 }), 100).map(c => c.length)).toEqual([100, 100, 50]);
    expect(copyNameFor('A')).toBe('A (copy)');
    expect(copyNameFor('x'.repeat(120))).toHaveLength(120);
  });
});
