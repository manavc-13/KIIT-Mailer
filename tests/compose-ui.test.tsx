// @vitest-environment jsdom
// Smoke test: the Compose page, the send progress page and the active-job banner render against a
// mocked API, step navigation works, and the job request carries the built recipients.
import { describe, it, expect, vi, beforeAll, afterEach } from 'vitest';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { MemoryRouter, Route, Routes } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import type { Job, SenderAccount } from '../shared/types';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const sender: SenderAccount = {
  id: 's1', email: 'iqac@kiit.ac.in', displayName: 'IQAC KIIT', replyTo: null, defaultTheme: 'meridian',
  signatureHtml: '<p>sig</p>', signatureDoc: null, dailyLimit: 2000, usedToday: 312, disabled: false,
};
const job: Job = {
  id: 'j1', userId: 'u1', senderAccountId: 's1', senderEmail: sender.email, templateId: null, mode: 'personal',
  subject: 'Hello', status: 'paused', totalMessages: 2, totalRecipients: 2, sentCount: 1, failedCount: 1, skippedCount: 0,
  createdAt: new Date().toISOString(), startedAt: new Date().toISOString(), finishedAt: null,
};

const posts: { path: string; body: unknown }[] = [];
vi.mock('@/lib/api', async (orig) => {
  const real = await orig<typeof import('@/lib/api')>();
  const get = async (path: string) => {
    if (path === '/auth/me') return { user: { id: 'u1', email: 'me@kiit.ac.in', name: 'Me', role: 'sender', mustChangePassword: false, disabled: false, lastLoginAt: null, createdAt: '' } };
    if (path === '/senders') return { senders: [sender] };
    if (path === '/templates') return { templates: [] };
    if (path === '/uploads/config') return { provider: 'local', publicImages: false, maxRawAttachmentBytes: 18e6, maxEncodedMessageBytes: 25 * 1024 * 1024, softWarnBytes: 10 * 1024 * 1024, mimeOverheadBytes: 98304, degradedReason: null };
    if (path === '/jobs/active') return { jobs: [job] };
    if (path.startsWith('/jobs/j1')) return { job, messages: [
      { id: 'm1', seq: 1, to: 'a@x.com', cc: null, bcc: null, recipientCount: 1, status: 'sent', attempts: 1, error: null, retryable: null, sentAt: new Date().toISOString() },
      { id: 'm2', seq: 2, to: 'b@x.com', cc: null, bcc: null, recipientCount: 1, status: 'failed', attempts: 3, error: 'Mailbox full', retryable: false, sentAt: null },
    ] };
    throw new real.ApiClientError(404, { error: 'not mocked ' + path });
  };
  const post = async (path: string, body?: unknown) => {
    posts.push({ path, body });
    if (path === '/render/preview') return { subject: '', html: '', text: '', sizeBytes: 0, lint: [], webFontsUsed: [] };
    if (path === '/jobs') return { job: { ...job, id: 'j2', status: 'queued' }, warnings: [] };
    return {};
  };
  return { ...real, api: { ...real.api, get, post } };
});

let root: Root | null = null;
let container: HTMLDivElement;

async function mount(ui: React.ReactNode, path: string) {
  const { AuthProvider } = await import('@/lib/auth');
  const { TooltipProvider } = await import('@/components/ui/tooltip');
  container = document.createElement('div');
  document.body.appendChild(container);
  const qc = new QueryClient({ defaultOptions: { queries: { retry: false } } });
  await act(async () => {
    root = createRoot(container);
    root.render(
      <QueryClientProvider client={qc}>
        <MemoryRouter initialEntries={[path]}>
          <AuthProvider><TooltipProvider>{ui}</TooltipProvider></AuthProvider>
        </MemoryRouter>
      </QueryClientProvider>,
    );
  });
  await flush();
}
const flush = async () => { for (let i = 0; i < 5; i++) await act(async () => { await new Promise(r => setTimeout(r, 20)); }); };
const byText = (t: string) => [...container.querySelectorAll('button, a')].find(b => b.textContent?.includes(t)) as HTMLElement | undefined;
const setInput = async (el: HTMLInputElement, value: string) => {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
  await act(async () => { setter.call(el, value); el.dispatchEvent(new Event('input', { bubbles: true })); });
};

beforeAll(() => {
  window.matchMedia = window.matchMedia || ((q: string) => ({
    matches: false, media: q, onchange: null, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {}, dispatchEvent: () => false,
  }) as unknown as MediaQueryList);
  window.scrollTo = () => {};
  (globalThis as { ResizeObserver?: unknown }).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
});
afterEach(async () => {
  await act(async () => root?.unmount());
  container.remove();
  root = null;
});

describe('ComposePage', () => {
  it('auto-selects the only sender, applies its theme, and walks the steps', async () => {
    const { default: ComposePage } = await import('@/pages/ComposePage');
    const { useComposeStore } = await import('@/compose/store');
    await mount(<ComposePage />, '/compose');

    expect(container.textContent).toContain('IQAC KIIT');
    expect(container.textContent).toContain('312 / 2,000 today');
    expect(useComposeStore.getState().senderAccountId).toBe('s1');
    expect(useComposeStore.getState().theme).toBe('meridian');

    // Next is blocked without a recipient
    await act(async () => byText('Next: Compose')!.click());
    expect(useComposeStore.getState().step).toBe('recipients');

    await setInput(container.querySelector<HTMLInputElement>('#single-email')!, 'dean@kiit.ac.in');
    await setInput(container.querySelector<HTMLInputElement>('#single-name')!, 'Dean');
    await act(async () => byText('Next: Compose')!.click());
    expect(useComposeStore.getState().step).toBe('compose');

    await setInput(container.querySelector<HTMLInputElement>('#subject')!, 'Meeting for {Name}');
    useComposeStore.getState().edit({ bodyHtml: '<p>Dear {Name}, please join.</p>' });
    await flush();
    await act(async () => byText('Next: Review')!.click());
    await flush();
    expect(useComposeStore.getState().step).toBe('review');
    expect(container.textContent).toContain('Deliverability checklist');
    expect(container.textContent).toContain('Send to 1 recipient');
    expect(container.textContent).toContain('remaining after this send');

    useComposeStore.getState().set({ globalCc: 'registrar@kiit.ac.in' });
    await flush();
    await act(async () => byText('Send to 1 recipient')!.click());
    await flush();
    const confirm = [...document.querySelectorAll('[role="alertdialog"] button')].find(b => b.textContent?.includes('Send to 1 recipient')) as HTMLElement;
    expect(document.body.textContent).toContain('Send this email now?');
    await act(async () => confirm.click());
    await flush();
    const req = posts.find(p => p.path === '/jobs')?.body as import('../shared/types').CreateJobRequest;
    expect(req).toBeTruthy();
    expect(req.senderAccountId).toBe('s1');
    expect(req.mode).toBe('personal');
    expect(req.subject).toBe('Meeting for {Name}');
    expect(req.recipients).toEqual([{ to: 'dean@kiit.ac.in', cc: 'registrar@kiit.ac.in', data: { Name: 'Dean', Email: 'dean@kiit.ac.in' } }]);
    expect(req.unsubscribe).toBe('auto');
  });

  it('shows the empty state when the user has no sender accounts', async () => {
    const apiMod = await import('@/lib/api');
    const realGet = apiMod.api.get;
    apiMod.api.get = (async (p: string) => (p === '/senders' ? { senders: [] } : realGet(p))) as typeof realGet;
    const { default: ComposePage } = await import('@/pages/ComposePage');
    await mount(<ComposePage />, '/compose');
    expect(container.textContent).toContain('Ask the administrator to give you access to a sender account');
    apiMod.api.get = realGet;
  });
});

describe('SendProgressPage + banner', () => {
  it('renders counters and the per-recipient table for a paused job', async () => {
    const { default: SendProgressPage } = await import('@/pages/SendProgressPage');
    await mount(<Routes><Route path="/send/:jobId" element={<SendProgressPage />} /></Routes>, '/send/j1');
    expect(container.textContent).toContain('Sending paused');
    expect(container.textContent).toContain('Mailbox full');
    expect(byText('Resume')).toBeTruthy();
    expect(byText('Retry 1 failed')).toBeTruthy();
  });

  it('banner links to the active job outside /send', async () => {
    const { ActiveJobBanner } = await import('@/components/ActiveJobBanner');
    await mount(<ActiveJobBanner />, '/compose');
    const link = container.querySelector('a[href="/send/j1"]');
    expect(link).toBeTruthy();
    expect(container.textContent).toContain('Hello');
  });
});
