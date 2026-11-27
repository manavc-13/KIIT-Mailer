// @vitest-environment jsdom
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { MemoryRouter, Route, Routes } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const toastFn = vi.hoisted(() => Object.assign(vi.fn(), { warning: vi.fn(), success: vi.fn(), error: vi.fn(), info: vi.fn() }));
vi.mock('sonner', () => ({ toast: toastFn }));

import { AuthProvider } from '@/lib/auth';
import { TooltipProvider } from '@/components/ui/tooltip';
import LoginPage from '@/pages/LoginPage';
import UsersPage from '@/pages/admin/UsersPage';
import { detectLegacyData, offerLegacyCleanup } from '@/lib/legacy';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const ADMIN = {
  id: 'u1', email: 'admin@kiit.ac.in', name: 'Asha Admin', role: 'admin', mustChangePassword: false,
  disabled: false, lastLoginAt: null, createdAt: '2026-01-01T00:00:00Z',
};

function json(status: number, body: unknown) {
  return Promise.resolve(new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } }));
}

let root: Root; let el: HTMLDivElement;
beforeEach(() => { el = document.createElement('div'); document.body.appendChild(el); root = createRoot(el); });
afterEach(() => { act(() => root.unmount()); el.remove(); vi.restoreAllMocks(); localStorage.clear(); });

async function render(ui: React.ReactNode, path: string) {
  const qc = new QueryClient({ defaultOptions: { queries: { retry: false } } });
  await act(async () => {
    root.render(
      <QueryClientProvider client={qc}>
        <MemoryRouter initialEntries={[path]}>
          <AuthProvider><TooltipProvider>{ui}</TooltipProvider></AuthProvider>
        </MemoryRouter>
      </QueryClientProvider>,
    );
  });
  await act(async () => { await new Promise(r => setTimeout(r, 20)); });
}

describe('LoginPage', () => {
  it('renders the form and shows the server error inline', async () => {
    const fetchMock = vi.fn((url: string) => {
      if (url.endsWith('/auth/me')) return json(401, { error: 'Please sign in.' });
      if (url.endsWith('/auth/login')) return json(401, { error: 'Incorrect email or password.' });
      return json(404, { error: 'nf' });
    });
    vi.stubGlobal('fetch', fetchMock);
    await render(<Routes><Route path="/login" element={<LoginPage />} /></Routes>, '/login');

    expect(el.textContent).toContain('Internal Quality Assurance Cell');
    expect(el.querySelector('label[for="login-email"]')).not.toBeNull();
    const email = el.querySelector<HTMLInputElement>('#login-email')!;
    const pw = el.querySelector<HTMLInputElement>('#login-password')!;
    expect(pw.type).toBe('password');
    await act(async () => { el.querySelector<HTMLButtonElement>('button[aria-label="Show password"]')!.click(); });
    expect(pw.type).toBe('text');

    const setValue = (input: HTMLInputElement, v: string) => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, v);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    };
    await act(async () => { setValue(email, 'someone@kiit.ac.in'); setValue(pw, 'wrong-password'); });
    await act(async () => { el.querySelector('form')!.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })); });
    await act(async () => { await new Promise(r => setTimeout(r, 20)); });

    const loginCall = fetchMock.mock.calls.find(c => String(c[0]).endsWith('/auth/login')) as unknown as [string, RequestInit];
    expect(loginCall).toBeTruthy();
    expect((loginCall[1].headers as Record<string, string>)['X-Requested-With']).toBe('kiit-mailer');
    expect(el.querySelector('[role="alert"]')?.textContent).toContain('Incorrect email or password.');
  });
});

describe('UsersPage', () => {
  it('lists users with role, status and allotted sender chips', async () => {
    vi.stubGlobal('fetch', vi.fn((url: string) => {
      if (url.endsWith('/auth/me')) return json(200, { user: ADMIN });
      if (url.endsWith('/admin/users')) return json(200, { users: [
        { ...ADMIN, senderAccountIds: [] },
        { ...ADMIN, id: 'u2', email: 'ravi@kiit.ac.in', name: 'Ravi Sender', role: 'sender', mustChangePassword: true, senderAccountIds: ['s1'] },
      ] });
      if (url.endsWith('/admin/senders')) return json(200, { senders: [
        { id: 's1', email: 'iqac@kiit.ac.in', displayName: 'IQAC Office', replyTo: null, defaultTheme: 'heritage', signatureHtml: null, signatureDoc: null,
          dailyLimit: 2000, usedToday: 0, disabled: false, appPasswordHint: 'mnop', lastVerifiedAt: null, lastVerifyError: null, createdAt: '2026-01-01T00:00:00Z', userIds: ['u2'] },
      ] });
      return json(404, { error: 'nf' });
    }));
    await render(<UsersPage />, '/admin/users');
    const text = el.textContent ?? '';
    expect(text).toContain('Ravi Sender');
    expect(text).toContain('Administrator');
    expect(text).toContain('Awaiting first sign-in');
    expect(text).toContain('IQAC Office');
    expect(text).toContain('(you)');
  });
});

describe('legacy v1 cleanup', () => {
  it('detects v1 keys and offers removal, leaving templates alone', () => {
    localStorage.setItem('kiit_mailer_settings', JSON.stringify({ email: 'x@kiit.ac.in', password: 'abcd efgh ijkl mnop' }));
    localStorage.setItem('kiit_mailer_logs', '[]');
    localStorage.setItem('kiit_mailer_templates', '[]');
    expect(detectLegacyData()).toEqual({ settings: true, leftovers: ['kiit_mailer_logs'] });

    offerLegacyCleanup();
    expect(toastFn.warning).toHaveBeenCalledTimes(1);
    const opts = toastFn.warning.mock.calls[0][1] as { action: { label: string; onClick: () => void } };
    expect(opts.action.label).toBe('Remove it');
    opts.action.onClick();
    expect(localStorage.getItem('kiit_mailer_settings')).toBeNull();

    const leftoverOpts = toastFn.mock.calls[0][1] as { action: { onClick: () => void } };
    leftoverOpts.action.onClick();
    expect(localStorage.getItem('kiit_mailer_logs')).toBeNull();
    expect(localStorage.getItem('kiit_mailer_templates')).toBe('[]');
  });
});
