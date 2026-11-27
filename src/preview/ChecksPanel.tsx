// Collapsible "Checks" panel under the preview: size vs Gmail's clip limit, content
// lint from the server, web fonts that fall back, and caniemail compatibility.
import { useMemo, useState } from 'react';
import { ChevronDown, CircleCheck, CircleX, ExternalLink, Info, TriangleAlert } from 'lucide-react';
import { cn } from '@/lib/utils';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import { findFont, FONT_CATALOG } from '../../shared/fonts';
import { formatBytes, GMAIL_CLIP_BYTES } from '../../shared/text';
import type { LintIssue, RenderResult } from '../../shared/types';
import { CANIEMAIL, compatibilityReport, type CompatRow, type SupportLevel } from './caniemail';
import { TARGET_CLIENTS } from './caniemail-spec';
import { hasOutlookConditionals } from './features';
import { readPref, writePref } from './prefs';

const LEVEL_META: Record<LintIssue['level'], { label: string; icon: typeof Info; cls: string }> = {
  error: { label: 'Must fix', icon: CircleX, cls: 'text-destructive' },
  warn: { label: 'Warnings', icon: TriangleAlert, cls: 'text-warning' },
  info: { label: 'Notes', icon: Info, cls: 'text-muted-foreground' },
};

const SUPPORT_META: Record<SupportLevel, { label: string; cls: string }> = {
  y: { label: 'Yes', cls: 'border-success/30 bg-success-soft text-success' },
  a: { label: 'Partial', cls: 'border-warning/30 bg-warning-soft text-warning' },
  n: { label: 'No', cls: 'border-destructive/30 bg-destructive/10 text-destructive' },
  u: { label: 'Untested', cls: 'border-border bg-muted text-muted-foreground' },
};

function SizeMeter({ bytes }: { bytes: number }) {
  const ratio = bytes / GMAIL_CLIP_BYTES;
  const tone = ratio > 1 ? 'bg-destructive' : ratio > 0.8 ? 'bg-warning' : 'bg-success';
  const pct = Math.min(100, ratio * 100);
  return (
    <section aria-labelledby="pv-size">
      <div className="flex items-baseline justify-between gap-3">
        <h4 id="pv-size" className="font-sans text-[13px] font-semibold">Message size</h4>
        <span className="text-[12.5px] tabular-nums text-muted-foreground">
          {formatBytes(bytes)} of {formatBytes(GMAIL_CLIP_BYTES)} Gmail limit
        </span>
      </div>
      <div
        role="meter"
        aria-label="HTML size against Gmail's clipping limit"
        aria-valuemin={0}
        aria-valuemax={GMAIL_CLIP_BYTES}
        aria-valuenow={bytes}
        aria-valuetext={`${formatBytes(bytes)} of ${formatBytes(GMAIL_CLIP_BYTES)}`}
        className="mt-2 h-2 overflow-hidden rounded-full bg-muted"
      >
        <div className={cn('h-full rounded-full', tone)} style={{ width: `${pct}%` }} />
      </div>
      <p className="mt-1.5 text-[12px] text-muted-foreground">
        {ratio > 1
          ? 'Gmail will clip this message and hide the rest behind "View entire message". Shorten it or move content to a link.'
          : ratio > 0.8
            ? 'Close to the limit where Gmail clips the message.'
            : 'Well under the limit where Gmail clips messages.'}
      </p>
    </section>
  );
}

function LintList({ lint }: { lint: LintIssue[] }) {
  if (!lint.length) {
    return (
      <p className="flex items-center gap-2 text-[13px] text-success">
        <CircleCheck className="size-4" /> No content issues found.
      </p>
    );
  }
  return (
    <div className="space-y-3">
      {(['error', 'warn', 'info'] as const).map(level => {
        const items = lint.filter(i => i.level === level);
        if (!items.length) return null;
        const meta = LEVEL_META[level];
        const Icon = meta.icon;
        return (
          <div key={level}>
            <div className={cn('mb-1 text-[11.5px] font-semibold uppercase tracking-wide', meta.cls)}>{meta.label} ({items.length})</div>
            <ul className="space-y-1">
              {items.map((i, n) => (
                <li key={i.code + n} className="flex items-start gap-2 text-[13px] leading-5">
                  <Icon className={cn('mt-0.5 size-4 shrink-0', meta.cls)} aria-hidden />
                  <span>{i.message}</span>
                </li>
              ))}
            </ul>
          </div>
        );
      })}
    </div>
  );
}

function FontFallbacks({ fonts }: { fonts: string[] }) {
  if (!fonts.length) {
    return <p className="text-[13px] text-muted-foreground">Only universal fonts are used, so every client shows the same typeface.</p>;
  }
  return (
    <ul className="space-y-1">
      {fonts.map(name => {
        const f = findFont(name) ?? FONT_CATALOG.find(x => x.id === name);
        const fallback = f?.fallback.split(',')[0].replace(/['"]/g, '').trim();
        return (
          <li key={name} className="flex flex-wrap items-baseline gap-x-2 text-[13px]">
            <span className="font-medium">{f?.label ?? name}</span>
            <span className="text-muted-foreground">shows as</span>
            <span className="font-medium">{fallback ?? 'a fallback font'}</span>
            <span className="text-[12px] text-muted-foreground">in Gmail and Outlook for Windows</span>
          </li>
        );
      })}
    </ul>
  );
}

function SupportChip({ row, clientId }: { row: CompatRow; clientId: (typeof TARGET_CLIENTS)[number]['id'] }) {
  const s = row.support[clientId];
  const client = TARGET_CLIENTS.find(c => c.id === clientId)!;
  const meta = SUPPORT_META[s.level];
  const chip = (
    <span
      tabIndex={s.notes.length ? 0 : undefined}
      className={cn('inline-flex items-center gap-1 rounded border px-1.5 py-px text-[11.5px] font-medium', meta.cls)}
    >
      {client.short}: {meta.label}
    </span>
  );
  if (!s.notes.length) return chip;
  return (
    <Tooltip>
      <TooltipTrigger asChild>{chip}</TooltipTrigger>
      <TooltipContent className="max-w-xs text-left">
        <div className="font-semibold">{client.label}</div>
        {s.notes.map((n, i) => <div key={i}>{n}</div>)}
      </TooltipContent>
    </Tooltip>
  );
}

function Compatibility({ html }: { html: string }) {
  const rows = useMemo(() => compatibilityReport(html), [html]);
  const problems = rows.filter(r => r.problems.length);
  const fine = rows.filter(r => !r.problems.length);
  const mso = useMemo(() => hasOutlookConditionals(html), [html]);
  return (
    <div>
      {mso && (
        <p className="mb-2 text-[12.5px] text-muted-foreground">
          This email includes Outlook-only fallback code, which may cover some of the Outlook gaps below.
        </p>
      )}
      {problems.length === 0 ? (
        <p className="flex items-center gap-2 text-[13px] text-success">
          <CircleCheck className="size-4" /> Everything detected is supported in Gmail, Apple Mail and Outlook.
        </p>
      ) : (
        <ul className="divide-y rounded-md border">
          {problems.map(r => (
            <li key={r.slug} className="px-3 py-2">
              <div className="flex flex-wrap items-baseline gap-x-2">
                <a href={r.url} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 font-mono text-[12.5px] font-medium text-foreground underline-offset-2 hover:underline">
                  {r.title} <ExternalLink className="size-3 opacity-60" aria-hidden /><span className="sr-only">(caniemail, opens in a new tab)</span>
                </a>
                <span className="text-[12px] text-muted-foreground">{r.evidence}</span>
              </div>
              <div className="mt-1.5 flex flex-wrap gap-1">
                {r.problems.map(id => <SupportChip key={id} row={r} clientId={id} />)}
              </div>
            </li>
          ))}
        </ul>
      )}
      {fine.length > 0 && (
        <p className="mt-2 text-[12px] text-muted-foreground">
          Also used, supported in all four: {fine.map(r => r.title).join(', ')}.
        </p>
      )}
      <p className="mt-2 text-[11.5px] text-muted-foreground">
        Support data from <a className="underline underline-offset-2" href="https://www.caniemail.com/" target="_blank" rel="noreferrer">caniemail.com</a>
        {' '}(CC BY 4.0), updated {CANIEMAIL.updated.slice(0, 10)}. Hover a chip for details.
      </p>
    </div>
  );
}

export function ChecksPanel({ result, className, defaultOpen = false }: { result: RenderResult; className?: string; defaultOpen?: boolean }) {
  const [open, setOpen] = useState(() => readPref('checksOpen', defaultOpen ? '1' : '0') === '1');
  const errors = result.lint.filter(l => l.level === 'error').length;
  const warns = result.lint.filter(l => l.level === 'warn').length;
  const compatProblems = useMemo(() => compatibilityReport(result.html).filter(r => r.problems.length).length, [result.html]);
  const toggle = () => { const next = !open; setOpen(next); writePref('checksOpen', next ? '1' : '0'); };

  const summary: { text: string; cls: string }[] = [];
  if (errors) summary.push({ text: `${errors} to fix`, cls: 'text-destructive' });
  if (warns) summary.push({ text: `${warns} warning${warns > 1 ? 's' : ''}`, cls: 'text-warning' });
  if (result.sizeBytes > GMAIL_CLIP_BYTES) summary.push({ text: 'clipped in Gmail', cls: 'text-destructive' });
  if (compatProblems) summary.push({ text: `${compatProblems} compatibility note${compatProblems > 1 ? 's' : ''}`, cls: 'text-muted-foreground' });
  if (!summary.length) summary.push({ text: 'all clear', cls: 'text-success' });

  return (
    <div className={cn('rounded-lg border bg-card', className)}>
      <button
        type="button"
        onClick={toggle}
        aria-expanded={open}
        aria-controls="pv-checks"
        className="flex w-full items-center gap-3 rounded-lg px-4 py-2.5 text-left hover:bg-muted/60 focus-visible:outline-2 focus-visible:outline-ring"
      >
        <span className="font-serif text-[15px] font-semibold">Checks</span>
        <span className="flex min-w-0 flex-1 flex-wrap gap-x-2 text-[12.5px]">
          {summary.map((s, i) => <span key={i} className={s.cls}>{i > 0 && <span className="mr-2 text-muted-foreground">·</span>}{s.text}</span>)}
        </span>
        <span className="text-[12px] tabular-nums text-muted-foreground">{formatBytes(result.sizeBytes)}</span>
        <ChevronDown className={cn('size-4 text-muted-foreground transition-transform', open && 'rotate-180')} aria-hidden />
      </button>
      {open && (
        <div id="pv-checks" className="space-y-5 border-t px-4 pb-4 pt-4">
          <SizeMeter bytes={result.sizeBytes} />
          <section aria-labelledby="pv-lint">
            <h4 id="pv-lint" className="mb-2 font-sans text-[13px] font-semibold">Content</h4>
            <LintList lint={result.lint} />
          </section>
          <section aria-labelledby="pv-fonts">
            <h4 id="pv-fonts" className="mb-2 font-sans text-[13px] font-semibold">Fonts that fall back in Gmail and Outlook</h4>
            <FontFallbacks fonts={result.webFontsUsed} />
          </section>
          <section aria-labelledby="pv-compat">
            <h4 id="pv-compat" className="mb-2 font-sans text-[13px] font-semibold">Client compatibility</h4>
            <Compatibility html={result.html} />
          </section>
        </div>
      )}
    </div>
  );
}
