import { describe, it, expect } from 'vitest';
import { computeMeter, admitFiles, DEFAULT_BUDGET, utf8Bytes } from '../src/compose/attachmentMath';

const MB = 1024 * 1024;

describe('computeMeter', () => {
  it('sizes against encoded bytes (4/3 base64) + body + MIME overhead', () => {
    const m = computeMeter(3 * MB, 1000);
    expect(m.encodedBytes).toBe(Math.ceil((3 * MB) / 3) * 4 + 1000 + DEFAULT_BUDGET.mimeOverheadBytes);
    expect(m.level).toBe('ok');
    expect(m.largeWarning).toBe(false);
    expect(m.percent).toBe(Math.round((m.encodedBytes / (25 * MB)) * 100));
  });
  it('warns above the soft limit and turns danger above the 25 MB encoded cap', () => {
    expect(computeMeter(8 * MB, 0).level).toBe('warn'); // ~10.7 MB encoded
    expect(computeMeter(11 * MB, 0).largeWarning).toBe(true);
    const over = computeMeter(19 * MB, 0);
    expect(over.level).toBe('danger');
    expect(over.percent).toBe(100);
  });
});

describe('admitFiles', () => {
  it('skips empty files, rejects files over the per-file cap and over the encoded budget', () => {
    const files = [
      { name: 'empty.txt', size: 0 },
      { name: 'huge.zip', size: 19 * MB },
      { name: 'a.pdf', size: 10 * MB },
      { name: 'b.pdf', size: 9 * MB }, // 19 MB total -> ~25.4 MB encoded: rejected
      { name: 'c.pdf', size: 1 * MB },
    ];
    const r = admitFiles(files, 0, 0);
    expect(r.accepted.map(f => f.name)).toEqual(['a.pdf', 'c.pdf']);
    expect(r.rejected.map(x => x.file.name)).toEqual(['empty.txt', 'huge.zip', 'b.pdf']);
  });
  it('counts already attached bytes', () => {
    expect(admitFiles([{ name: 'x', size: 2 * MB }], 17 * MB, 0).accepted).toHaveLength(0);
  });
  it('utf8Bytes counts multibyte text', () => {
    expect(utf8Bytes('ab', 'é', null)).toBe(4);
  });
});
