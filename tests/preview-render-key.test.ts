// renderRequestKey ties a render result to the exact request it came from (Review uses it to
// ignore results for an older sender / unsubscribe / recipient).
import { describe, it, expect } from 'vitest';
import { renderRequestKey } from '@/preview/useRenderPreview';
import type { RenderRequest } from '../shared/types';

const base = {
  senderAccountId: 's1', subject: 'Hi {Name}', bodyHtml: '<p>Hello</p>', theme: 'heritage', includeSignature: true,
  row: { Name: 'Dean', Email: 'dean@kiit.ac.in' },
} as unknown as RenderRequest;

describe('renderRequestKey', () => {
  it('is stable for equal requests and null for none', () => {
    expect(renderRequestKey({ ...base })).toBe(renderRequestKey({ ...base }));
    expect(renderRequestKey(null)).toBeNull();
  });
  it('changes with the sender, unsubscribe and recipient row', () => {
    const k = renderRequestKey(base);
    expect(renderRequestKey({ ...base, senderAccountId: 's2' })).not.toBe(k);
    expect(renderRequestKey({ ...base, unsubscribe: true } as RenderRequest)).not.toBe(k);
    expect(renderRequestKey({ ...base, row: { Name: 'Registrar', Email: 'reg@kiit.ac.in' } } as RenderRequest)).not.toBe(k);
  });
});
