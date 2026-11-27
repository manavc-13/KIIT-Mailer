// Small building blocks shared by the admin screens.
import { useState, type ReactNode } from 'react';
import { Check, Copy, KeyRound, Search } from 'lucide-react';
import { toast } from 'sonner';
import { cn } from '@/lib/utils';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Skeleton } from '@/components/ui/skeleton';
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription,
  AlertDialogFooter, AlertDialogHeader, AlertDialogTitle,
} from '@/components/ui/alert-dialog';

/** Centered message for an empty list, with an optional call to action. */
export function EmptyState({ icon: Icon, title, children, action, className }: {
  icon: typeof KeyRound; title: string; children?: ReactNode; action?: ReactNode; className?: string;
}) {
  return (
    <div className={cn('flex flex-col items-center rounded-lg border border-dashed bg-card px-6 py-14 text-center', className)}>
      <span className="grid size-11 place-items-center rounded-full bg-gold-soft text-[var(--gold)]">
        <Icon className="size-5" strokeWidth={1.75} />
      </span>
      <h2 className="mt-4 text-lg font-semibold">{title}</h2>
      {children && <div className="mt-1 max-w-md text-sm text-muted-foreground">{children}</div>}
      {action && <div className="mt-5">{action}</div>}
    </div>
  );
}

/** Confirmation dialog for consequential actions. `onConfirm` may be async; the dialog stays open while it runs. */
export function ConfirmDialog({
  open, onOpenChange, title, description, confirmLabel, destructive, onConfirm,
}: {
  open: boolean; onOpenChange: (open: boolean) => void; title: string; description: ReactNode;
  confirmLabel: string; destructive?: boolean; onConfirm: () => Promise<unknown> | void;
}) {
  const [busy, setBusy] = useState(false);
  return (
    <AlertDialog open={open} onOpenChange={(o) => { if (!busy) onOpenChange(o); }}>
      <AlertDialogContent>
        <AlertDialogHeader>
          <AlertDialogTitle className="font-serif">{title}</AlertDialogTitle>
          <AlertDialogDescription asChild><div className="space-y-2 text-sm text-muted-foreground">{description}</div></AlertDialogDescription>
        </AlertDialogHeader>
        <AlertDialogFooter>
          <AlertDialogCancel disabled={busy}>Cancel</AlertDialogCancel>
          <AlertDialogAction
            variant={destructive ? 'destructive' : 'default'}
            disabled={busy}
            onClick={async (e) => {
              e.preventDefault();
              setBusy(true);
              try { await onConfirm(); onOpenChange(false); } catch { /* caller shows the error */ } finally { setBusy(false); }
            }}
          >
            {busy ? 'Working…' : confirmLabel}
          </AlertDialogAction>
        </AlertDialogFooter>
      </AlertDialogContent>
    </AlertDialog>
  );
}

export function CopyButton({ value, label = 'Copy', className }: { value: string; label?: string; className?: string }) {
  const [copied, setCopied] = useState(false);
  return (
    <Button
      type="button" variant="outline" size="sm" className={className}
      onClick={async () => {
        try {
          await navigator.clipboard.writeText(value);
          setCopied(true);
          setTimeout(() => setCopied(false), 2000);
        } catch {
          toast.error('Could not copy automatically. Select the text and copy it manually.');
        }
      }}
    >
      {copied ? <Check /> : <Copy />} {copied ? 'Copied' : label}
    </Button>
  );
}

/** One-time display of a temporary password, with copy and a clear handling note. */
export function TempPasswordReveal({ password, who }: { password: string; who: string }) {
  return (
    <div className="space-y-3">
      <div className="rounded-md border border-[var(--gold)]/40 bg-gold-soft p-4">
        <div className="text-xs font-medium uppercase tracking-[0.12em] text-accent-foreground">Temporary password for {who}</div>
        <div className="mt-2 flex flex-wrap items-center justify-between gap-3">
          <code className="select-all font-mono text-xl tracking-wider text-foreground" aria-label="Temporary password">{password}</code>
          <CopyButton value={password} />
        </div>
      </div>
      <p className="text-sm text-muted-foreground">
        Share this password with the user securely, in person or over a direct message, not in a group or by email to a shared inbox.
        <strong className="font-medium text-foreground"> It will not be shown again.</strong> They will be asked to choose their own password the first time they sign in.
      </p>
    </div>
  );
}

export function SearchBox({ value, onChange, placeholder, label, className }: {
  value: string; onChange: (v: string) => void; placeholder: string; label: string; className?: string;
}) {
  return (
    <div className={cn('relative w-full max-w-xs', className)}>
      <Search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" aria-hidden="true" />
      <Input type="search" aria-label={label} placeholder={placeholder} value={value} onChange={e => onChange(e.target.value)} className="bg-card pl-8" />
    </div>
  );
}

export function TableSkeleton({ rows = 5, cols = 5 }: { rows?: number; cols?: number }) {
  return (
    <div className="space-y-2 rounded-lg border bg-card p-4" aria-busy="true" aria-label="Loading">
      {Array.from({ length: rows }).map((_, i) => (
        <div key={i} className="flex gap-4">
          {Array.from({ length: cols }).map((__, j) => <Skeleton key={j} className="h-5 flex-1" />)}
        </div>
      ))}
    </div>
  );
}

/** Small coloured status label. */
export function Pill({ tone = 'neutral', children, className, title }: {
  tone?: 'neutral' | 'success' | 'warning' | 'danger' | 'gold'; children: ReactNode; className?: string; title?: string;
}) {
  const tones = {
    neutral: 'border-border bg-secondary text-secondary-foreground',
    success: 'border-[var(--success)]/25 bg-success-soft text-success',
    warning: 'border-[var(--warning)]/25 bg-warning-soft text-warning',
    danger: 'border-destructive/25 bg-destructive/10 text-destructive',
    gold: 'border-[var(--gold)]/40 bg-gold-soft text-accent-foreground',
  } as const;
  return (
    <span title={title} className={cn('inline-flex items-center gap-1 whitespace-nowrap rounded-full border px-2 py-0.5 text-xs font-medium [&>svg]:size-3', tones[tone], className)}>
      {children}
    </span>
  );
}
