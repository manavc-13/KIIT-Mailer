// @vitest-environment jsdom
// Send progress pump: one pump per mount under StrictMode, an idle wait when /run makes no progress,
// and no "Retry N failed" while this tab is sending.
import { describe, it, expect, vi, beforeAll, afterEach } from 'vitest';
import { act, StrictMode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { MemoryRouter, Route, Routes } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import type { Job } from '../shared/types';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const job: Job = {
  id: 'j1', userId: 'u1', senderAccountId: 's1', senderEmail: 'iqac@kiit.ac.in', templateId: null, mode: 'personal',
  subject: 'Hello', status: 'sending', totalMessages: 3, totalRecipients: 3, sentCount: 1, failedCount: 1, skippedCount: 0,
  createdAt: new Date().toISOString(), startedAt: new Date().toISOString(), finishedAt: null,
};

const runCalls: number[] = [];
vi.mock('@/lib/api', async (orig) => {
  const real = await orig<typeof import('@/lib/api')>();
  const get = async (path: string) => {
    if (path === '/auth/me') return { user: { id: 'u1', email: 'me@kiit.ac.in', name: 'Me', role: 'sender', mustChangePassword: false, disabled: false, lastLoginAt: null, createdAt: '' } };
    if (path.startsWith('/jobs/j1')) return { job, messages: [] };
    throw new real.ApiClientError(404, { error: 'not mocked ' + path });
  };
  const post = async (path: string) => {
    if (path.endsWith('/run')) {
      runCalls.push(Date.now());
      // Another tab holds the claim: not done, nothing processed.
      return { job, processed: [], done: false };
    }
    return {};
  };
  return { ...real, api: { ...real.api, get, post } };
});

let root: Root | null = null;
let container: HTMLDivElement;
const flush = async (rounds = 5) => { for (let i = 0; i < rounds; i++) await act(async () => { await new Promise(r => setTimeout(r, 20)); }); };
const byText = (t: string) => [...container.querySelectorAll('button, a')].find(b => b.textContent?.includes(t)) as HTMLElement | undefined;

beforeAll(() => {
  window.matchMedia = window.matchMedia || ((q: string) => ({
    matches: false, media: q, onchange: null, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {}, dispatchEvent: () => false,
  }) as unknown as MediaQueryList);
});
afterEach(async () => {
  await act(async () => root?.unmount());
  container.remove();
  root = null;
});

describe('SendProgressPage pump', () => {
  it('runs a single pump under StrictMode and waits after a no-progress /run', async () => {
    const { AuthProvider } = await import('@/lib/auth');
    const { default: SendProgressPage } = await import('@/pages/SendProgressPage');
    container = document.createElement('div');
    document.body.appendChild(container);
    const qc = new QueryClient({ defaultOptions: { queries: { retry: false } } });
    await act(async () => {
      root = createRoot(container);
      root.render(
        <StrictMode>
          <QueryClientProvider client={qc}>
            <MemoryRouter initialEntries={['/send/j1']}>
              <AuthProvider><Routes><Route path="/send/:jobId" element={<SendProgressPage />} /></Routes></AuthProvider>
            </MemoryRouter>
          </QueryClientProvider>
        </StrictMode>,
      );
    });
    await flush(15); // ~300 ms: well under the 5 s idle wait

    expect(runCalls.length).toBe(1);
    expect(container.textContent).toContain('Sending from this tab');
    // Failed messages exist, but retrying is not offered while this tab is sending.
    expect(byText('Retry 1 failed')).toBeUndefined();
    expect(byText('Cancel send')).toBeTruthy();
  });
});

describe('api error mapping', () => {
  it('maps HTTP 413 to a plain-English message and titles quota/auth codes', async () => {
    const real = await vi.importActual<typeof import('@/lib/api')>('@/lib/api');
    const fetchMock = vi.fn(async () => new Response('Request Entity Too Large', { status: 413, headers: { 'content-type': 'text/plain' } }));
    vi.stubGlobal('fetch', fetchMock);
    try {
      await expect(real.api.post('/jobs', {})).rejects.toMatchObject({ status: 413, message: real.TOO_LARGE_MESSAGE });
    } finally {
      vi.unstubAllGlobals();
    }
    const quota = new real.ApiClientError(400, { error: 'Today’s limit is reached.', code: 'quota_exhausted' });
    expect(real.sendErrorTitle(quota, 'Could not start sending')).toBe('Daily sending limit reached');
    expect(real.errorMessage(quota)).toBe('Today’s limit is reached.');
    expect(real.sendErrorTitle(new real.ApiClientError(400, { error: 'x', code: 'auth_failed' }), 'f')).toMatch(/App Password/);
    expect(real.sendErrorTitle(new real.ApiClientError(400, { error: 'x' }), 'Could not start sending')).toBe('Could not start sending');
  });
});
