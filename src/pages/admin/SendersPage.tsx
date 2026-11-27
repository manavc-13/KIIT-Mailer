// Admin: sender accounts. Each is a @kiit.ac.in Google Workspace mailbox used via
// a Google App Password. The password is tested against Gmail before it is saved,
// stored encrypted on the server, and never shown again (only "••••" + last 4).
import { useMemo, useState, type FormEvent, type ReactNode } from 'react';
import { useMutation, useQueryClient } from '@tanstack/react-query';
import {
  AlertTriangle, CheckCircle2, ExternalLink, KeyRound, Loader2, Mail, Pencil, Plus, RefreshCw, Users, XCircle,
} from 'lucide-react';
import { toast } from 'sonner';
import { PageHeader, PageBody } from '@/components/AppShell';
import { api, errorMessage } from '@/lib/api';
import { cn } from '@/lib/utils';
import { EmailEditor, type EmailEditorValue } from '@/editor';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Switch } from '@/components/ui/switch';
import { Separator } from '@/components/ui/separator';
import { Skeleton } from '@/components/ui/skeleton';
import {
  Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle,
} from '@/components/ui/dialog';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { EmptyState, Pill } from '@/admin/components';
import { ADMIN_SENDERS_KEY, useAdminSenders, useAdminUsers } from '@/admin/queries';
import { formatDateTime, relativeTime } from '@/admin/format';
import { themeLabel, themeOptions } from '@/admin/themes';
import type {
  AdminUser, CreateSenderRequest, SenderAccountAdmin, ThemeId, UpdateSenderRequest,
} from '../../../shared/types';

const KIIT_SUFFIX = '@kiit.ac.in';
const cleanAppPassword = (v: string) => v.replace(/\s+/g, '');
const appPasswordShapeOk = (v: string) => /^[a-zA-Z]{16}$/.test(cleanAppPassword(v));

// ------------------------------------------------------------------ shared bits

function Field({ id, label, hint, children, className }: { id: string; label: string; hint?: ReactNode; children: ReactNode; className?: string }) {
  return (
    <div className={cn('space-y-2', className)}>
      <Label htmlFor={id}>{label}</Label>
      {children}
      {hint && <p id={`${id}-hint`} className="text-xs text-muted-foreground">{hint}</p>}
    </div>
  );
}

function ThemeSelect({ id, value, onChange }: { id: string; value: ThemeId; onChange: (v: ThemeId) => void }) {
  return (
    <Select value={value} onValueChange={v => onChange(v as ThemeId)}>
      <SelectTrigger id={id} className="w-full bg-card"><SelectValue /></SelectTrigger>
      <SelectContent>
        {themeOptions().map(t => (
          <SelectItem key={t.id} value={t.id}>
            <span className="flex items-center gap-2">
              <span className="flex overflow-hidden rounded-sm border" aria-hidden="true">
                {t.swatch.slice(0, 3).map((c, i) => <span key={i} className="block size-3" style={{ background: c }} />)}
              </span>
              {t.label}
              <span className="text-xs text-muted-foreground">· {t.description}</span>
            </span>
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}

function AppPasswordHelp() {
  return (
    <details className="group rounded-md border bg-muted/40 px-4 py-3 text-sm open:pb-4">
      <summary className="cursor-pointer select-none font-medium marker:text-[var(--gold)]">
        How to create a Google App Password for this account
      </summary>
      <ol className="mt-3 list-decimal space-y-2 pl-5 text-muted-foreground">
        <li>Sign in to the <strong className="font-medium text-foreground">sender account itself</strong> (not your personal account) in a browser.</li>
        <li>
          Turn on <strong className="font-medium text-foreground">2-Step Verification</strong>: Google Account → Security → 2-Step Verification.
          App Passwords only exist once this is on. If the option is missing, ask KIIT IT to allow it for the account.
        </li>
        <li>
          Open{' '}
          <a href="https://myaccount.google.com/apppasswords" target="_blank" rel="noreferrer noopener" className="inline-flex items-center gap-1 font-medium text-foreground underline underline-offset-2">
            myaccount.google.com/apppasswords <ExternalLink className="size-3" aria-hidden="true" />
          </a>
          , enter a name such as “KIIT Mailer”, and choose Create.
        </li>
        <li>Google shows a 16-letter password in four groups (for example <code className="rounded bg-card px-1 font-mono text-xs">abcd efgh ijkl mnop</code>). Copy it and paste it below; spaces are fine.</li>
        <li>The password is shown by Google only once. If it is lost, delete it there and create a new one.</li>
      </ol>
    </details>
  );
}

type TestState = { status: 'idle' } | { status: 'testing' } | { status: 'ok'; key: string } | { status: 'error'; key: string; error: string };

/** "Test connection" button + result line. Reports success for an (email, password) key. */
function useConnectionTest() {
  const [state, setState] = useState<TestState>({ status: 'idle' });
  async function run(email: string, appPassword: string) {
    const key = `${email.trim().toLowerCase()}|${cleanAppPassword(appPassword)}`;
    setState({ status: 'testing' });
    try {
      const res = await api.post<{ ok: boolean; error?: string }>('/admin/senders/test-credentials', { email: email.trim(), appPassword });
      setState(res.ok ? { status: 'ok', key } : { status: 'error', key, error: res.error || 'The connection test failed.' });
    } catch (err) {
      setState({ status: 'error', key, error: errorMessage(err) });
    }
  }
  const passedFor = (email: string, appPassword: string) =>
    state.status === 'ok' && state.key === `${email.trim().toLowerCase()}|${cleanAppPassword(appPassword)}`;
  return { state, run, passedFor, reset: () => setState({ status: 'idle' }) };
}

function TestResult({ state, stale }: { state: TestState; stale: boolean }) {
  if (state.status === 'idle') return null;
  if (state.status === 'testing') {
    return <p className="flex items-center gap-2 text-sm text-muted-foreground" aria-live="polite"><Loader2 className="size-4 animate-spin" /> Signing in to smtp.gmail.com… this can take up to 15 seconds.</p>;
  }
  if (stale) return <p className="text-sm text-muted-foreground" aria-live="polite">The details changed since the last test. Test the connection again.</p>;
  if (state.status === 'ok') {
    return <p className="flex items-center gap-2 text-sm font-medium text-success" aria-live="polite"><CheckCircle2 className="size-4" /> Connected. Gmail accepted this App Password.</p>;
  }
  return (
    <p role="alert" className="flex gap-2 rounded-md border border-destructive/30 bg-destructive/5 px-3 py-2 text-sm text-destructive">
      <XCircle className="mt-0.5 size-4 shrink-0" /> <span>{state.error}</span>
    </p>
  );
}

// ------------------------------------------------------------------ add dialog

function AddSenderDialog({ open, onOpenChange }: { open: boolean; onOpenChange: (o: boolean) => void }) {
  const qc = useQueryClient();
  const [email, setEmail] = useState('');
  const [displayName, setDisplayName] = useState('');
  const [replyTo, setReplyTo] = useState('');
  const [appPassword, setAppPassword] = useState('');
  const [dailyLimit, setDailyLimit] = useState('2000');
  const [theme, setTheme] = useState<ThemeId>('heritage');
  const [error, setError] = useState<string | null>(null);
  const test = useConnectionTest();

  const emailOk = email.trim().toLowerCase().endsWith(KIIT_SUFFIX) && /^\S+@\S+$/.test(email.trim());
  const pwOk = appPasswordShapeOk(appPassword);
  const passed = test.passedFor(email, appPassword);
  const stale = (test.state.status === 'ok' || test.state.status === 'error') && !passed
    && test.state.key !== `${email.trim().toLowerCase()}|${cleanAppPassword(appPassword)}`;

  const reset = () => {
    setEmail(''); setDisplayName(''); setReplyTo(''); setAppPassword(''); setDailyLimit('2000'); setTheme('heritage');
    setError(null); test.reset();
  };
  /** Every way out (Cancel, Esc, outside click, ×, or a successful save) closes and clears the form, App Password included. */
  const close = () => {
    onOpenChange(false);
    setTimeout(reset, 200); // after the close animation
  };

  const create = useMutation({
    mutationFn: (req: CreateSenderRequest) => api.post<{ sender: SenderAccountAdmin }>('/admin/senders', req),
    onSuccess: (res) => {
      qc.invalidateQueries({ queryKey: ADMIN_SENDERS_KEY });
      qc.invalidateQueries({ queryKey: ['senders'] });
      toast.success(`${res.sender.displayName} was added and verified. Allot it to users on the Users page.`);
      close();
    },
    onError: (err) => setError(errorMessage(err)),
  });

  function submit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    const limit = Number(dailyLimit);
    if (!displayName.trim()) return setError('Enter a display name. Recipients see it as the sender’s name.');
    if (!Number.isInteger(limit) || limit < 1 || limit > 2000) return setError('The daily limit must be a whole number from 1 to 2,000.');
    if (!passed) return setError('Test the connection first. The account can be saved once Gmail accepts the App Password.');
    create.mutate({
      email: email.trim().toLowerCase(), displayName: displayName.trim(), replyTo: replyTo.trim() || null,
      appPassword, dailyLimit: limit, defaultTheme: theme,
    });
  }

  return (
    <Dialog open={open} onOpenChange={(o) => { if (create.isPending) return; if (!o) close(); else onOpenChange(true); }}>
      <DialogContent className="max-h-[92vh] overflow-y-auto sm:max-w-2xl">
        <form onSubmit={submit} noValidate className="grid gap-5">
          <DialogHeader>
            <DialogTitle className="font-serif text-xl">Add a sender account</DialogTitle>
            <DialogDescription>
              Connect a @kiit.ac.in mailbox. The App Password is checked with Gmail, then stored encrypted on the server. Nobody, including administrators, can view it afterwards.
            </DialogDescription>
          </DialogHeader>
          {error && <p role="alert" className="rounded-md border border-destructive/30 bg-destructive/5 px-3 py-2 text-sm text-destructive">{error}</p>}

          <div className="grid gap-4 sm:grid-cols-2">
            <Field id="as-email" label="Account email" hint={email && !emailOk ? 'Sender accounts must be @kiit.ac.in addresses.' : 'The Google Workspace mailbox mail is sent from.'}>
              <Input id="as-email" type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="iqac.office@kiit.ac.in" autoFocus aria-invalid={!!email && !emailOk || undefined} aria-describedby="as-email-hint" />
            </Field>
            <Field id="as-name" label="Display name" hint="Shown to recipients, e.g. “IQAC, KIIT”.">
              <Input id="as-name" value={displayName} onChange={e => setDisplayName(e.target.value)} maxLength={120} aria-describedby="as-name-hint" />
            </Field>
            <Field id="as-reply" label="Reply-to (optional)" hint="Replies go here instead of the account itself.">
              <Input id="as-reply" type="email" value={replyTo} onChange={e => setReplyTo(e.target.value)} aria-describedby="as-reply-hint" />
            </Field>
            <Field id="as-limit" label="Daily recipient limit" hint="Google Workspace allows up to 2,000 recipients a day per account.">
              <Input id="as-limit" type="number" min={1} max={2000} value={dailyLimit} onChange={e => setDailyLimit(e.target.value)} aria-describedby="as-limit-hint" />
            </Field>
            <Field id="as-theme" label="Default theme" className="sm:col-span-2">
              <ThemeSelect id="as-theme" value={theme} onChange={setTheme} />
            </Field>
          </div>

          <Separator />

          <div className="space-y-3">
            <AppPasswordHelp />
            <Field id="as-pw" label="Google App Password" hint={appPassword && !pwOk ? 'An App Password is 16 letters. Spaces are removed automatically.' : undefined}>
              <div className="flex flex-col gap-2 sm:flex-row">
                <Input
                  id="as-pw" type="password" autoComplete="off" spellCheck={false} value={appPassword}
                  onChange={e => setAppPassword(e.target.value)} placeholder="abcd efgh ijkl mnop" className="font-mono"
                  aria-describedby="as-pw-hint"
                />
                <Button
                  type="button" variant="outline" className="shrink-0"
                  disabled={!emailOk || !pwOk || test.state.status === 'testing'}
                  onClick={() => test.run(email, appPassword)}
                >
                  {test.state.status === 'testing' ? <Loader2 className="animate-spin" /> : <RefreshCw />} Test connection
                </Button>
              </div>
            </Field>
            <TestResult state={test.state} stale={stale} />
          </div>

          <DialogFooter>
            <Button type="button" variant="ghost" onClick={close} disabled={create.isPending}>Cancel</Button>
            <Button type="submit" disabled={!passed || create.isPending} title={!passed ? 'Test the connection first' : undefined}>
              {create.isPending ? <><Loader2 className="animate-spin" /> Saving…</> : 'Save account'}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

// ------------------------------------------------------------------ edit dialog

function EditSenderDialog({ sender, onClose }: { sender: SenderAccountAdmin; onClose: () => void }) {
  const qc = useQueryClient();
  const [displayName, setDisplayName] = useState(sender.displayName);
  const [replyTo, setReplyTo] = useState(sender.replyTo ?? '');
  const [dailyLimit, setDailyLimit] = useState(String(sender.dailyLimit));
  const [theme, setTheme] = useState<ThemeId>(sender.defaultTheme);
  const [enabled, setEnabled] = useState(!sender.disabled);
  const [newPassword, setNewPassword] = useState('');
  const [signature, setSignature] = useState<EmailEditorValue | null>(null);
  // Stable initial values for the editor (it owns the document while open).
  const [initialSig] = useState(() => ({ doc: sender.signatureDoc ?? null, html: sender.signatureHtml ?? '' }));
  const [error, setError] = useState<string | null>(null);
  const test = useConnectionTest();

  const rotating = newPassword.trim().length > 0;
  const rotationPassed = test.passedFor(sender.email, newPassword);
  const stale = (test.state.status === 'ok' || test.state.status === 'error') && !rotationPassed
    && test.state.key !== `${sender.email}|${cleanAppPassword(newPassword)}`;

  const save = useMutation({
    mutationFn: (req: UpdateSenderRequest) => api.patch<{ sender: SenderAccountAdmin }>(`/admin/senders/${sender.id}`, req),
    onSuccess: (res, req) => {
      qc.invalidateQueries({ queryKey: ADMIN_SENDERS_KEY });
      qc.invalidateQueries({ queryKey: ['senders'] }); // prefix: includes ['senders', 'mine'] (Compose's From list)
      // Signature, theme, display name and reply-to all feed the rendered email: re-render open previews.
      qc.invalidateQueries({ queryKey: ['render-preview'] });
      toast.success(req.appPassword ? `Saved ${res.sender.displayName}. The new App Password is in use.` : `Saved ${res.sender.displayName}.`);
      onClose();
    },
    onError: (err) => setError(errorMessage(err)),
  });

  function submit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    const limit = Number(dailyLimit);
    if (!displayName.trim()) return setError('Enter a display name.');
    if (!Number.isInteger(limit) || limit < 1 || limit > 2000) return setError('The daily limit must be a whole number from 1 to 2,000.');
    if (rotating && !rotationPassed) return setError('Test the new App Password before saving, or clear the field to keep the current one.');

    const req: UpdateSenderRequest = {};
    if (displayName.trim() !== sender.displayName) req.displayName = displayName.trim();
    if ((replyTo.trim() || null) !== sender.replyTo) req.replyTo = replyTo.trim() || null;
    if (limit !== sender.dailyLimit) req.dailyLimit = limit;
    if (theme !== sender.defaultTheme) req.defaultTheme = theme;
    if (enabled === sender.disabled) req.disabled = !enabled;
    if (rotating) req.appPassword = newPassword;
    if (signature) {
      const empty = !signature.html.replace(/<[^>]*>|&nbsp;|\s/g, '');
      req.signatureHtml = empty ? null : signature.html;
      req.signatureDoc = empty ? null : signature.doc;
    }
    if (!Object.keys(req).length) return onClose();
    save.mutate(req);
  }

  return (
    <Dialog open onOpenChange={(o) => { if (!o && !save.isPending) onClose(); }}>
      <DialogContent className="max-h-[92vh] overflow-y-auto sm:max-w-3xl">
        <form onSubmit={submit} noValidate className="grid gap-5">
          <DialogHeader>
            <DialogTitle className="font-serif text-xl">Edit {sender.displayName}</DialogTitle>
            <DialogDescription>{sender.email}</DialogDescription>
          </DialogHeader>
          {error && <p role="alert" className="rounded-md border border-destructive/30 bg-destructive/5 px-3 py-2 text-sm text-destructive">{error}</p>}

          <Tabs defaultValue="details">
            <TabsList>
              <TabsTrigger value="details">Details</TabsTrigger>
              <TabsTrigger value="signature">Signature</TabsTrigger>
              <TabsTrigger value="password">App Password</TabsTrigger>
            </TabsList>

            <TabsContent value="details" className="mt-4 grid gap-4 sm:grid-cols-2">
              <Field id="es-name" label="Display name">
                <Input id="es-name" value={displayName} onChange={e => setDisplayName(e.target.value)} maxLength={120} />
              </Field>
              <Field id="es-reply" label="Reply-to (optional)">
                <Input id="es-reply" type="email" value={replyTo} onChange={e => setReplyTo(e.target.value)} />
              </Field>
              <Field id="es-limit" label="Daily recipient limit" hint={`${sender.usedToday.toLocaleString('en-IN')} used today. Maximum 2,000.`}>
                <Input id="es-limit" type="number" min={1} max={2000} value={dailyLimit} onChange={e => setDailyLimit(e.target.value)} aria-describedby="es-limit-hint" />
              </Field>
              <Field id="es-theme" label="Default theme" hint="Applied when someone picks this account in Compose.">
                <ThemeSelect id="es-theme" value={theme} onChange={setTheme} />
              </Field>
              <div className="flex items-start justify-between gap-4 rounded-md border bg-card p-4 sm:col-span-2">
                <div>
                  <Label htmlFor="es-enabled" className="text-sm font-medium">Account enabled</Label>
                  <p className="mt-1 text-xs text-muted-foreground">
                    A disabled account disappears from everyone’s From list and cannot send. Allotments are kept.
                  </p>
                </div>
                <Switch id="es-enabled" checked={enabled} onCheckedChange={setEnabled} />
              </div>
            </TabsContent>

            <TabsContent value="signature" className="mt-4 space-y-2">
              <p className="text-sm text-muted-foreground">
                Added below the message body when this account sends, unless the sender switches it off for a message.
              </p>
              <div className="rounded-md border bg-card">
                <EmailEditor
                  variant="signature"
                  doc={initialSig.doc}
                  initialHtml={initialSig.html}
                  theme={theme}
                  minHeight={180}
                  onChange={setSignature}
                />
              </div>
            </TabsContent>

            <TabsContent value="password" className="mt-4 space-y-3">
              <div className="flex items-center gap-2 text-sm">
                <KeyRound className="size-4 text-muted-foreground" aria-hidden="true" />
                Current App Password: <code className="font-mono">••••{sender.appPasswordHint ?? ''}</code>
              </div>
              <p className="text-sm text-muted-foreground">
                To replace it (for example after it was revoked in Google), paste a new App Password and test it. It takes effect when you save.
              </p>
              <AppPasswordHelp />
              <Field id="es-pw" label="New App Password">
                <div className="flex flex-col gap-2 sm:flex-row">
                  <Input
                    id="es-pw" type="password" autoComplete="off" spellCheck={false} value={newPassword}
                    onChange={e => setNewPassword(e.target.value)} placeholder="Leave empty to keep the current one" className="font-mono"
                  />
                  <Button
                    type="button" variant="outline" className="shrink-0"
                    disabled={!appPasswordShapeOk(newPassword) || test.state.status === 'testing'}
                    onClick={() => test.run(sender.email, newPassword)}
                  >
                    {test.state.status === 'testing' ? <Loader2 className="animate-spin" /> : <RefreshCw />} Test connection
                  </Button>
                </div>
              </Field>
              {rotating && <TestResult state={test.state} stale={stale} />}
            </TabsContent>
          </Tabs>

          <DialogFooter>
            <Button type="button" variant="ghost" onClick={onClose} disabled={save.isPending}>Cancel</Button>
            <Button type="submit" disabled={save.isPending || (rotating && !rotationPassed)}>
              {save.isPending ? <><Loader2 className="animate-spin" /> Saving…</> : 'Save changes'}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

// ------------------------------------------------------------------ card

function VerifyStatus({ sender }: { sender: SenderAccountAdmin }) {
  if (sender.lastVerifyError) {
    return (
      <div role="status" className="flex gap-2 rounded-md border border-destructive/30 bg-destructive/5 px-3 py-2 text-sm text-destructive">
        <AlertTriangle className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
        <span><span className="font-medium">Last check failed.</span> {sender.lastVerifyError}</span>
      </div>
    );
  }
  if (sender.lastVerifiedAt) {
    return (
      <p className="flex items-center gap-2 text-sm text-success" title={formatDateTime(sender.lastVerifiedAt)}>
        <CheckCircle2 className="size-4" aria-hidden="true" /> Verified {relativeTime(sender.lastVerifiedAt)}
      </p>
    );
  }
  return <p className="text-sm text-muted-foreground">Not verified yet.</p>;
}

function SenderCard({ sender, users, onEdit }: { sender: SenderAccountAdmin; users: Map<string, AdminUser>; onEdit: () => void }) {
  const qc = useQueryClient();
  const verify = useMutation({
    mutationFn: () => api.post<{ ok: boolean; error?: string; sender: SenderAccountAdmin }>(`/admin/senders/${sender.id}/verify`),
    onSuccess: (res) => {
      qc.invalidateQueries({ queryKey: ADMIN_SENDERS_KEY });
      if (res.ok) toast.success(`${sender.displayName}: Gmail accepted the stored App Password.`);
      else toast.error(`${sender.displayName}: ${res.error}`);
    },
    onError: (err) => toast.error(errorMessage(err)),
  });
  const allotted = sender.userIds.map(id => users.get(id)).filter((u): u is AdminUser => !!u);
  const pct = Math.min(100, Math.round((sender.usedToday / Math.max(1, sender.dailyLimit)) * 100));

  return (
    <article className={cn('flex flex-col rounded-lg border bg-card', sender.disabled && 'opacity-75')} aria-labelledby={`sender-${sender.id}`}>
      <header className="flex items-start justify-between gap-3 border-b px-5 py-4">
        <div className="min-w-0">
          <h2 id={`sender-${sender.id}`} className="truncate text-lg font-semibold">{sender.displayName}</h2>
          <p className="flex items-center gap-1.5 truncate text-sm text-muted-foreground"><Mail className="size-3.5" aria-hidden="true" /> {sender.email}</p>
        </div>
        {sender.disabled ? <Pill tone="danger">Disabled</Pill> : <Pill tone="success">Enabled</Pill>}
      </header>

      <div className="flex-1 space-y-4 px-5 py-4">
        <VerifyStatus sender={sender} />
        <dl className="grid grid-cols-2 gap-x-6 gap-y-3 text-sm">
          <div>
            <dt className="text-xs uppercase tracking-wide text-muted-foreground">App Password</dt>
            <dd className="mt-0.5 font-mono">••••{sender.appPasswordHint ?? ''}</dd>
          </div>
          <div>
            <dt className="text-xs uppercase tracking-wide text-muted-foreground">Default theme</dt>
            <dd className="mt-0.5">{themeLabel(sender.defaultTheme)}</dd>
          </div>
          <div>
            <dt className="text-xs uppercase tracking-wide text-muted-foreground">Reply-to</dt>
            <dd className="mt-0.5 truncate">{sender.replyTo ?? <span className="text-muted-foreground">Same as account</span>}</dd>
          </div>
          <div>
            <dt className="text-xs uppercase tracking-wide text-muted-foreground">Signature</dt>
            <dd className="mt-0.5">{sender.signatureHtml ? 'Set' : <span className="text-muted-foreground">None</span>}</dd>
          </div>
          <div className="col-span-2">
            <dt className="text-xs uppercase tracking-wide text-muted-foreground">Sent today</dt>
            <dd className="mt-1">
              <div className="flex items-baseline justify-between text-sm">
                <span>{sender.usedToday.toLocaleString('en-IN')} of {sender.dailyLimit.toLocaleString('en-IN')} recipients</span>
                <span className="text-xs text-muted-foreground">{pct}%</span>
              </div>
              <div className="mt-1 h-1.5 overflow-hidden rounded-full bg-muted" role="presentation">
                <div className={cn('h-full rounded-full', pct >= 90 ? 'bg-destructive' : 'bg-[var(--gold)]')} style={{ width: `${pct}%` }} />
              </div>
            </dd>
          </div>
          <div className="col-span-2">
            <dt className="flex items-center gap-1.5 text-xs uppercase tracking-wide text-muted-foreground"><Users className="size-3.5" aria-hidden="true" /> Allotted to</dt>
            <dd className="mt-1.5 flex flex-wrap gap-1">
              {allotted.length === 0
                ? <span className="text-muted-foreground">No users yet (administrators can always use it)</span>
                : allotted.map(u => (
                  <span key={u.id} className={cn('rounded border bg-secondary px-1.5 py-0.5 text-xs text-secondary-foreground', u.disabled && 'line-through opacity-70')} title={u.email}>
                    {u.name}
                  </span>
                ))}
            </dd>
          </div>
        </dl>
      </div>

      <footer className="flex flex-wrap items-center justify-end gap-2 border-t bg-muted/30 px-5 py-3">
        <Button variant="outline" size="sm" onClick={() => verify.mutate()} disabled={verify.isPending}>
          {verify.isPending ? <Loader2 className="animate-spin" /> : <RefreshCw />} {verify.isPending ? 'Checking…' : 'Re-verify'}
        </Button>
        <Button size="sm" onClick={onEdit}><Pencil /> Edit</Button>
      </footer>
    </article>
  );
}

// ------------------------------------------------------------------ page

export default function SendersPage() {
  const sendersQ = useAdminSenders();
  const usersQ = useAdminUsers();
  const users = useMemo(() => new Map((usersQ.data ?? []).map(u => [u.id, u])), [usersQ.data]);
  const [addOpen, setAddOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const editing = sendersQ.data?.find(s => s.id === editingId) ?? null;

  return (
    <>
      <PageHeader
        title="Sender accounts"
        description="The @kiit.ac.in mailboxes that official mail is sent from. App Passwords are stored encrypted and can be replaced but never viewed."
        actions={<Button onClick={() => setAddOpen(true)}><Plus /> Add account</Button>}
      />
      <PageBody>
        {sendersQ.isLoading ? (
          <div className="grid gap-5 lg:grid-cols-2">
            {[0, 1].map(i => <Skeleton key={i} className="h-80 rounded-lg" />)}
          </div>
        ) : sendersQ.isError ? (
          <p role="alert" className="text-sm text-destructive">Could not load sender accounts: {errorMessage(sendersQ.error)}</p>
        ) : !sendersQ.data?.length ? (
          <EmptyState icon={KeyRound} title="No sender accounts yet" action={<Button onClick={() => setAddOpen(true)}><Plus /> Add the first account</Button>}>
            Add a @kiit.ac.in mailbox with a Google App Password. Users can send only from accounts you allot to them.
          </EmptyState>
        ) : (
          <div className="grid gap-5 lg:grid-cols-2">
            {sendersQ.data.map(s => <SenderCard key={s.id} sender={s} users={users} onEdit={() => setEditingId(s.id)} />)}
          </div>
        )}
      </PageBody>

      <AddSenderDialog open={addOpen} onOpenChange={setAddOpen} />
      {editing && <EditSenderDialog key={editing.id} sender={editing} onClose={() => setEditingId(null)} />}
    </>
  );
}

