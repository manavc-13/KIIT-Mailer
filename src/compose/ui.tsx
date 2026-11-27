// Small presentational helpers used across the compose flow. Solid surfaces, subtle borders.
import { useRef, type ReactNode, type KeyboardEvent } from 'react';
import type { LucideIcon } from 'lucide-react';
import { cn } from '@/lib/utils';

export function Section({ title, description, actions, children, className, id }: {
  title: ReactNode; description?: ReactNode; actions?: ReactNode; children: ReactNode; className?: string; id?: string;
}) {
  return (
    <section id={id} className={cn('rounded-lg border bg-card shadow-xs', className)} aria-labelledby={id ? `${id}-title` : undefined}>
      <header className="flex flex-wrap items-start justify-between gap-3 border-b px-5 py-3.5">
        <div className="min-w-0">
          <h2 id={id ? `${id}-title` : undefined} className="text-[17px] font-semibold leading-snug">{title}</h2>
          {description && <p className="mt-0.5 text-[13px] text-muted-foreground">{description}</p>}
        </div>
        {actions && <div className="flex flex-wrap items-center gap-2">{actions}</div>}
      </header>
      <div className="px-5 py-4">{children}</div>
    </section>
  );
}

export interface SegmentOption<T extends string> { value: T; label: ReactNode; icon?: LucideIcon; hint?: string }

/** Accessible segmented control (radiogroup with roving focus and arrow keys). */
export function Segmented<T extends string>({ value, onChange, options, ariaLabel, size = 'md', className }: {
  value: T; onChange: (v: T) => void; options: SegmentOption<T>[]; ariaLabel: string; size?: 'sm' | 'md'; className?: string;
}) {
  const refs = useRef<(HTMLButtonElement | null)[]>([]);
  const onKey = (e: KeyboardEvent, i: number) => {
    const dir = e.key === 'ArrowRight' || e.key === 'ArrowDown' ? 1 : e.key === 'ArrowLeft' || e.key === 'ArrowUp' ? -1 : 0;
    if (!dir) return;
    e.preventDefault();
    const next = (i + dir + options.length) % options.length;
    onChange(options[next].value);
    refs.current[next]?.focus();
  };
  return (
    <div role="radiogroup" aria-label={ariaLabel} className={cn('inline-flex rounded-md border bg-muted p-0.5', className)}>
      {options.map((o, i) => {
        const active = o.value === value;
        const Icon = o.icon;
        return (
          <button
            key={o.value}
            ref={el => { refs.current[i] = el; }}
            type="button"
            role="radio"
            aria-checked={active}
            tabIndex={active ? 0 : -1}
            title={o.hint}
            onClick={() => onChange(o.value)}
            onKeyDown={e => onKey(e, i)}
            className={cn(
              'inline-flex items-center gap-1.5 rounded-[5px] font-medium transition-colors outline-none focus-visible:ring-2 focus-visible:ring-ring',
              size === 'sm' ? 'px-2.5 py-1 text-xs' : 'px-3.5 py-1.5 text-[13px]',
              active ? 'bg-card text-foreground shadow-xs ring-1 ring-border' : 'text-muted-foreground hover:text-foreground',
            )}
          >
            {Icon && <Icon className="size-3.5" strokeWidth={1.75} />}
            {o.label}
          </button>
        );
      })}
    </div>
  );
}

export type Tone = 'ok' | 'warn' | 'danger' | 'neutral' | 'info';

export const toneClasses: Record<Tone, string> = {
  ok: 'bg-success-soft text-success border-success/25',
  warn: 'bg-warning-soft text-warning border-warning/25',
  danger: 'bg-destructive/10 text-destructive border-destructive/25',
  neutral: 'bg-muted text-muted-foreground border-border',
  info: 'bg-gold-soft text-accent-foreground border-gold/30',
};

export function StatPill({ label, value, tone }: { label: string; value: ReactNode; tone: Tone }) {
  return (
    <div className={cn('flex min-w-[8.5rem] flex-1 flex-col rounded-md border px-3 py-2', toneClasses[tone])}>
      <span className="text-[11px] font-medium uppercase tracking-wide opacity-80">{label}</span>
      <span className="text-lg font-semibold tabular-nums">{value}</span>
    </div>
  );
}

export function Chip({ children, tone = 'neutral', className }: { children: ReactNode; tone?: Tone; className?: string }) {
  return (
    <span className={cn('inline-flex max-w-full items-center gap-1 truncate rounded border px-1.5 py-0.5 text-xs font-medium', toneClasses[tone], className)}>
      {children}
    </span>
  );
}

/** Character counter: muted, then warning, then danger. */
export function CharCounter({ count, warnAbove, dangerAbove, warnBelow, id }: {
  count: number; warnAbove: number; dangerAbove?: number; warnBelow?: number; id?: string;
}) {
  const tone = dangerAbove !== undefined && count > dangerAbove ? 'text-destructive'
    : count > warnAbove || (warnBelow !== undefined && count > 0 && count < warnBelow) ? 'text-warning'
      : 'text-muted-foreground';
  return <span id={id} className={cn('text-xs tabular-nums', tone)} aria-live="polite">{count} chars</span>;
}

/** Thin solid meter bar. */
export function Meter({ percent, tone, label, className }: { percent: number; tone: 'ok' | 'warn' | 'danger' | 'primary'; label: string; className?: string }) {
  const fill = tone === 'danger' ? 'bg-destructive' : tone === 'warn' ? 'bg-warning' : tone === 'ok' ? 'bg-success' : 'bg-primary';
  return (
    <div
      role="meter" aria-label={label} aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(percent)}
      className={cn('h-1.5 w-full overflow-hidden rounded-full bg-muted ring-1 ring-inset ring-border', className)}
    >
      <div className={cn('h-full rounded-full transition-[width]', fill)} style={{ width: `${Math.max(0, Math.min(100, percent))}%` }} />
    </div>
  );
}

export function downloadText(content: string, filename: string, mime = 'text/plain') {
  const blob = new Blob([content], { type: `${mime};charset=utf-8` });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function pluralize(n: number, one: string, many = one + 's'): string {
  return `${n.toLocaleString()} ${n === 1 ? one : many}`;
}
