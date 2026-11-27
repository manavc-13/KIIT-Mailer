// Small presentational pieces shared by the History list and Job detail pages.
import type { ReactNode } from 'react';
import { cn } from '@/lib/utils';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import type { Job, JobMode, JobStatus, MessageStatus } from '../../shared/types';
import { absoluteTime, pendingCount, relativeTime } from './format';

const pill = 'inline-flex items-center gap-1.5 whitespace-nowrap rounded-full border px-2 py-0.5 text-[11.5px] font-medium';

const JOB_STATUS: Record<JobStatus, { label: string; cls: string; dot: string }> = {
  queued: { label: 'Queued', cls: 'border-border bg-muted text-muted-foreground', dot: 'bg-muted-foreground' },
  sending: { label: 'Sending', cls: 'border-[#9db4cf] bg-[#e8eff7] text-[#1d4a7a] dark:border-[#2d4a6a] dark:bg-[#132233] dark:text-[#9cc0e6]', dot: 'bg-[#1d4a7a] dark:bg-[#9cc0e6] animate-pulse' },
  paused: { label: 'Paused', cls: 'border-warning/40 bg-warning-soft text-warning', dot: 'bg-warning' },
  completed: { label: 'Completed', cls: 'border-success/35 bg-success-soft text-success', dot: 'bg-success' },
  cancelled: { label: 'Cancelled', cls: 'border-destructive/30 bg-destructive/8 text-destructive', dot: 'bg-destructive' },
};

export function JobStatusBadge({ status, className }: { status: JobStatus; className?: string }) {
  const m = JOB_STATUS[status] ?? JOB_STATUS.queued;
  return <span className={cn(pill, m.cls, className)}><span className={cn('size-1.5 rounded-full', m.dot)} aria-hidden />{m.label}</span>;
}

const MSG_STATUS: Record<MessageStatus, { label: string; cls: string }> = {
  queued: { label: 'Pending', cls: 'border-border bg-muted text-muted-foreground' },
  sending: { label: 'Sending', cls: 'border-[#9db4cf] bg-[#e8eff7] text-[#1d4a7a] dark:border-[#2d4a6a] dark:bg-[#132233] dark:text-[#9cc0e6]' },
  sent: { label: 'Sent', cls: 'border-success/35 bg-success-soft text-success' },
  failed: { label: 'Failed', cls: 'border-destructive/30 bg-destructive/8 text-destructive' },
  skipped: { label: 'Skipped', cls: 'border-warning/40 bg-warning-soft text-warning' },
};

export function MessageStatusBadge({ status }: { status: MessageStatus }) {
  const m = MSG_STATUS[status] ?? MSG_STATUS.queued;
  return <span className={cn(pill, m.cls)}>{m.label}</span>;
}

export function ModeBadge({ mode }: { mode: JobMode }) {
  return (
    <span className={cn(pill, 'border-gold/40 bg-gold-soft text-accent-foreground')}>
      {mode === 'bcc_broadcast' ? 'BCC broadcast' : 'Personalised'}
    </span>
  );
}

/** Stacked bar: sent (green), failed (red), skipped (amber), pending (track). */
export function CountsBar({ job, className }: { job: Job; className?: string }) {
  const total = Math.max(1, job.totalMessages);
  const seg = (n: number) => `${(n / total) * 100}%`;
  const pending = pendingCount(job);
  const label = `${job.sentCount} sent, ${job.failedCount} failed, ${job.skippedCount} skipped${pending ? `, ${pending} pending` : ''} of ${job.totalMessages} messages`;
  return (
    <div className={cn('flex h-1.5 w-full overflow-hidden rounded-full bg-muted', className)} role="img" aria-label={label}>
      <span className="h-full bg-success" style={{ width: seg(job.sentCount) }} />
      <span className="h-full bg-destructive" style={{ width: seg(job.failedCount) }} />
      <span className="h-full bg-warning" style={{ width: seg(job.skippedCount) }} />
    </div>
  );
}

export function CountsInline({ job }: { job: Job }) {
  return (
    <span className="flex flex-wrap gap-x-2.5 text-[12px] tabular-nums">
      <span className="text-success"><b className="font-semibold">{job.sentCount}</b> sent</span>
      {job.failedCount > 0 && <span className="text-destructive"><b className="font-semibold">{job.failedCount}</b> failed</span>}
      {job.skippedCount > 0 && <span className="text-warning"><b className="font-semibold">{job.skippedCount}</b> skipped</span>}
      <span className="text-muted-foreground">of {job.totalMessages}</span>
    </span>
  );
}

export function TimeAgo({ iso, className }: { iso: string | null | undefined; className?: string }) {
  if (!iso) return <span className={className}>—</span>;
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <time dateTime={iso} tabIndex={0} className={cn('whitespace-nowrap underline decoration-dotted decoration-muted-foreground/40 underline-offset-4', className)}>
          {relativeTime(iso)}
        </time>
      </TooltipTrigger>
      <TooltipContent>{absoluteTime(iso)}</TooltipContent>
    </Tooltip>
  );
}

export function EmptyState({ icon, title, children, action }: { icon: ReactNode; title: string; children?: ReactNode; action?: ReactNode }) {
  return (
    <div className="flex flex-col items-center rounded-lg border border-dashed bg-card px-6 py-14 text-center">
      <div className="grid size-12 place-items-center rounded-full bg-gold-soft text-accent-foreground">{icon}</div>
      <h2 className="mt-4 text-[19px] font-semibold">{title}</h2>
      {children && <div className="mt-1 max-w-md text-[13.5px] text-muted-foreground">{children}</div>}
      {action && <div className="mt-5">{action}</div>}
    </div>
  );
}
