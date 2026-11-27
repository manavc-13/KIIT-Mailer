// Change password. Two modes:
//  - required: the user signed in with a temporary password and must set their own (no way out but sign-out);
//  - voluntary: reached from the account menu.
import { useMemo, useState, type FormEvent } from 'react';
import { useNavigate } from 'react-router-dom';
import { AlertCircle, Check, Circle, Eye, EyeOff, Loader2, ShieldCheck } from 'lucide-react';
import { toast } from 'sonner';
import { PageHeader, PageBody } from '@/components/AppShell';
import { api, errorMessage } from '@/lib/api';
import { useAuth } from '@/lib/auth';
import { cn } from '@/lib/utils';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';

const MIN_LENGTH = 10;

/** Mirrors validatePasswordPolicy in server/lib/passwords.ts (the server is authoritative). */
function policyChecks(pw: string, email: string) {
  const local = (email.split('@')[0] || '').toLowerCase();
  const lower = pw.toLowerCase();
  const classes = [/[a-z]/, /[A-Z]/, /[0-9]/, /[^a-zA-Z0-9]/].filter(re => re.test(pw)).length;
  return [
    { id: 'length', label: `At least ${MIN_LENGTH} characters`, ok: pw.length >= MIN_LENGTH },
    { id: 'variety', label: 'A mix of at least two of: lowercase, uppercase, digits, symbols', ok: classes >= 2 },
    {
      id: 'email', label: 'Does not contain your email address or the part before the @',
      ok: pw.length > 0 && lower !== email.toLowerCase() && !(local.length >= 3 && lower.includes(local)),
    },
  ];
}

function PasswordField({ id, label, value, onChange, autoComplete, describedBy, autoFocus }: {
  id: string; label: string; value: string; onChange: (v: string) => void; autoComplete: string;
  describedBy?: string; autoFocus?: boolean;
}) {
  const [show, setShow] = useState(false);
  return (
    <div className="space-y-2">
      <Label htmlFor={id}>{label}</Label>
      <div className="relative">
        <Input
          id={id} type={show ? 'text' : 'password'} autoComplete={autoComplete} value={value} autoFocus={autoFocus}
          onChange={e => onChange(e.target.value)} aria-describedby={describedBy} className="h-10 bg-card pr-11"
        />
        <button
          type="button" onClick={() => setShow(s => !s)}
          aria-label={show ? `Hide ${label.toLowerCase()}` : `Show ${label.toLowerCase()}`} aria-pressed={show} aria-controls={id}
          className="absolute right-1 top-1/2 grid size-8 -translate-y-1/2 place-items-center rounded text-muted-foreground hover:text-foreground focus-visible:outline-2 focus-visible:outline-[var(--ring)]"
        >
          {show ? <EyeOff className="size-4" /> : <Eye className="size-4" />}
        </button>
      </div>
    </div>
  );
}

export default function ChangePasswordPage() {
  const { user, refresh } = useAuth();
  const navigate = useNavigate();
  const required = !!user?.mustChangePassword;
  const email = user?.email ?? '';

  const [current, setCurrent] = useState('');
  const [next, setNext] = useState('');
  const [confirm, setConfirm] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const checks = useMemo(() => policyChecks(next, email), [next, email]);
  const policyOk = checks.every(c => c.ok);
  const matches = confirm.length > 0 && confirm === next;
  const differs = next.length === 0 || next !== current;
  const canSubmit = current.length > 0 && policyOk && matches && differs && !submitting;

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    if (!canSubmit) {
      if (!current) setError(required ? 'Enter the temporary password you signed in with.' : 'Enter your current password.');
      else if (!policyOk) setError('Your new password does not meet the requirements listed below.');
      else if (!matches) setError('The two new passwords do not match.');
      else if (!differs) setError('Choose a new password that is different from your current one.');
      return;
    }
    setSubmitting(true);
    setError(null);
    try {
      await api.post('/auth/change-password', { currentPassword: current, newPassword: next });
      await refresh();
      toast.success(required ? 'Your password is set. Welcome to KIIT Mailer.' : 'Your password has been changed. Other devices have been signed out.');
      navigate('/compose', { replace: true });
    } catch (err) {
      setError(errorMessage(err));
      setSubmitting(false);
    }
  }

  return (
    <>
      <PageHeader
        title={required ? 'Choose your password' : 'Change password'}
        description={required
          ? 'You signed in with a temporary password issued by the administrator. Choose your own password to continue.'
          : 'Changing your password signs you out on every other device.'}
      />
      <PageBody>
        <div className="grid max-w-4xl gap-8 lg:grid-cols-[minmax(0,1fr)_minmax(0,18rem)]">
          <form onSubmit={onSubmit} noValidate className="space-y-5 rounded-lg border bg-card p-6">
            {required && (
              <div className="flex gap-3 rounded-md border border-[var(--gold)]/40 bg-gold-soft px-4 py-3 text-sm text-accent-foreground">
                <ShieldCheck className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
                <p>Temporary passwords work only once. Until you set your own password, the rest of KIIT Mailer stays locked.</p>
              </div>
            )}
            {error && (
              <div role="alert" className="flex gap-2.5 rounded-md border border-destructive/30 bg-destructive/5 px-3.5 py-3 text-sm text-destructive">
                <AlertCircle className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
                <span>{error}</span>
              </div>
            )}

            {/* Hidden username field helps password managers save the right account. */}
            <input type="email" autoComplete="username" value={email} readOnly hidden />

            <PasswordField
              id="cp-current" label={required ? 'Temporary password' : 'Current password'} autoComplete="current-password"
              value={current} onChange={setCurrent} autoFocus
            />
            <PasswordField
              id="cp-new" label="New password" autoComplete="new-password" value={next} onChange={setNext} describedBy="cp-rules"
            />
            <div className="space-y-2">
              <PasswordField id="cp-confirm" label="Confirm new password" autoComplete="new-password" value={confirm} onChange={setConfirm} />
              {confirm.length > 0 && !matches && (
                <p className="text-xs text-destructive">The passwords do not match yet.</p>
              )}
              {next.length > 0 && !differs && (
                <p className="text-xs text-destructive">The new password must be different from the current one.</p>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-3 pt-1">
              <Button type="submit" disabled={submitting} className="min-w-40">
                {submitting ? <><Loader2 className="animate-spin" /> Saving…</> : required ? 'Set password and continue' : 'Change password'}
              </Button>
              {!required && (
                <Button type="button" variant="ghost" onClick={() => navigate(-1)} disabled={submitting}>Cancel</Button>
              )}
            </div>
          </form>

          <aside className="space-y-3">
            <h2 className="text-base font-semibold">Password requirements</h2>
            <ul id="cp-rules" className="space-y-2.5 text-sm" aria-live="polite">
              {checks.map(c => (
                <li key={c.id} className={cn('flex gap-2.5', c.ok ? 'text-success' : 'text-muted-foreground')}>
                  {c.ok
                    ? <Check className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
                    : <Circle className="mt-0.5 size-4 shrink-0 opacity-60" aria-hidden="true" />}
                  <span>{c.label}<span className="sr-only">{c.ok ? ' (met)' : ' (not met yet)'}</span></span>
                </li>
              ))}
            </ul>
            <p className="pt-2 text-xs leading-relaxed text-muted-foreground">
              A short phrase of three or four unrelated words, with a number, is both strong and easy to remember.
            </p>
          </aside>
        </div>
      </PageBody>
    </>
  );
}
