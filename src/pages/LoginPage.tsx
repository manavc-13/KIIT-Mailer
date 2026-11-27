// Sign-in screen. Accounts are issued by the administrator; there is no sign-up.
import { useState, type FormEvent } from 'react';
import { Navigate, useLocation, useNavigate } from 'react-router-dom';
import { AlertCircle, Eye, EyeOff, Loader2 } from 'lucide-react';
import { api, ApiClientError, errorMessage } from '@/lib/api';
import { useAuth } from '@/lib/auth';
import { useLegacyCleanup } from '@/lib/legacy';
import { Monogram } from '@/components/AppShell';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import type { MeResponse, PublicUser } from '../../shared/types';

function destinationFor(user: PublicUser, from: unknown): string {
  if (user.mustChangePassword) return '/change-password';
  if (typeof from === 'string' && from.startsWith('/') && !from.startsWith('//') && !from.startsWith('/login')) return from;
  return '/compose';
}

export default function LoginPage() {
  const { user, loading, setUser } = useAuth();
  const location = useLocation();
  const navigate = useNavigate();
  const offerLegacyCleanup = useLegacyCleanup();
  const from = (location.state as { from?: unknown } | null)?.from;

  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  if (!loading && user && !submitting) return <Navigate to={destinationFor(user, from)} replace />;

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    if (!email.trim() || !password) {
      setError('Enter your email address and password.');
      return;
    }
    setSubmitting(true);
    setError(null);
    try {
      const res = await api.post<MeResponse>('/auth/login', { email: email.trim(), password });
      setUser(res.user);
      offerLegacyCleanup();
      navigate(destinationFor(res.user, from), { replace: true });
    } catch (err) {
      setPassword('');
      if (err instanceof ApiClientError) setError(err.message);
      else if (err instanceof TypeError) setError('Could not reach the server. Check your connection and try again.');
      else setError(errorMessage(err));
      setSubmitting(false);
    }
  }

  return (
    <div className="grid min-h-full lg:grid-cols-[minmax(0,5fr)_minmax(0,6fr)]">
      {/* Institutional panel */}
      <aside className="relative flex flex-col justify-between bg-sidebar px-8 py-8 text-sidebar-foreground sm:px-12 lg:py-12">
        <div className="flex items-center gap-3">
          <Monogram className="size-11 rounded-md ring-1 ring-white/10" />
          <div className="leading-tight">
            <div className="font-serif text-xl font-semibold text-white">KIIT Mailer</div>
            <div className="text-[11px] uppercase tracking-[0.16em] text-sidebar-muted">Official mail service</div>
          </div>
        </div>

        <div className="my-10 max-w-md lg:my-0">
          <div className="text-[11px] font-semibold uppercase tracking-[0.2em] text-[var(--gold)]">Internal Quality Assurance Cell</div>
          <h1 className="mt-4 font-serif text-3xl font-semibold leading-snug text-white sm:text-[2.35rem]">
            KIIT Mailer
          </h1>
          <div className="mt-5 flex flex-col gap-[3px]" aria-hidden="true">
            <span className="block h-px w-24 bg-[var(--gold)]" />
            <span className="block h-px w-16 bg-[var(--gold)]/60" />
          </div>
          <p className="mt-6 text-[15px] leading-relaxed text-sidebar-foreground/85">
            The institutional mailing service for official correspondence from the Internal Quality Assurance Cell.
          </p>
          <p className="mt-3 text-sm leading-relaxed text-sidebar-muted">
            Access is by invitation only. Accounts are created by the KIIT Mailer administrator.
          </p>
        </div>

        <div className="hidden text-xs text-sidebar-muted lg:block">
          Kalinga Institute of Industrial Technology, Bhubaneswar
        </div>
      </aside>

      {/* Sign-in form */}
      <main className="flex items-center justify-center bg-background px-6 py-12 sm:px-10">
        <div className="w-full max-w-sm">
          <h2 className="font-serif text-2xl font-semibold">Sign in</h2>
          <p className="mt-1.5 text-sm text-muted-foreground">
            Use the email address and password given to you by the administrator.
          </p>

          <form onSubmit={onSubmit} noValidate className="mt-8 space-y-5" aria-describedby={error ? 'login-error' : undefined}>
            {error && (
              <div
                id="login-error" role="alert"
                className="flex gap-2.5 rounded-md border border-destructive/30 bg-destructive/5 px-3.5 py-3 text-sm text-destructive"
              >
                <AlertCircle className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
                <span>{error}</span>
              </div>
            )}

            <div className="space-y-2">
              <Label htmlFor="login-email">Email address</Label>
              <Input
                id="login-email" type="email" autoComplete="username" inputMode="email" autoFocus required
                placeholder="name@kiit.ac.in"
                value={email} onChange={e => setEmail(e.target.value)}
                aria-invalid={!!error || undefined}
                className="h-10 bg-card"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="login-password">Password</Label>
              <div className="relative">
                <Input
                  id="login-password" type={showPassword ? 'text' : 'password'} autoComplete="current-password" required
                  value={password} onChange={e => setPassword(e.target.value)}
                  aria-invalid={!!error || undefined}
                  className="h-10 bg-card pr-11"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(s => !s)}
                  aria-label={showPassword ? 'Hide password' : 'Show password'}
                  aria-pressed={showPassword}
                  aria-controls="login-password"
                  className="absolute right-1 top-1/2 grid size-8 -translate-y-1/2 place-items-center rounded text-muted-foreground hover:text-foreground focus-visible:outline-2 focus-visible:outline-[var(--ring)]"
                >
                  {showPassword ? <EyeOff className="size-4" /> : <Eye className="size-4" />}
                </button>
              </div>
            </div>

            <Button type="submit" className="h-10 w-full" disabled={submitting}>
              {submitting ? <><Loader2 className="animate-spin" /> Signing in…</> : 'Sign in'}
            </Button>
          </form>

          <p className="mt-8 border-t pt-5 text-xs leading-relaxed text-muted-foreground">
            Forgotten your password, or need access? Contact the KIIT Mailer administrator at the IQAC office.
            They can reset your password or create an account for you.
          </p>
        </div>
      </main>
    </div>
  );
}
