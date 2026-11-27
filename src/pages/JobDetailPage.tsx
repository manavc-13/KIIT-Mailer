// One send job: summary, actions (live view, retry failed, export, cancel) and the
// per-recipient message table.
import { Fragment, useEffect, useMemo, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query';
import { toast } from 'sonner';
import {
  ArrowLeft, Ban, ChevronDown, ChevronLeft, ChevronRight, Download, Loader2, RadioTower, RotateCcw, Search, SearchX,
} from 'lucide-react';
import { PageBody, PageHeader } from '@/components/AppShell';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Skeleton } from '@/components/ui/skeleton';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter,
  AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger,
} from '@/components/ui/alert-dialog';
import { api, ApiClientError, errorMessage } from '@/lib/api';
import { cn } from '@/lib/utils';
import type { Job, JobMessage } from '../../shared/types';
import { parseEmailList } from '../../shared/text';
import { CountsBar, EmptyState, JobStatusBadge, MessageStatusBadge, ModeBadge, TimeAgo } from '@/history/components';
import {
  absoluteTime, countAddresses, isActive, matchesFilter, matchesSearch, pendingCount, shouldPoll, type MessageFilter,
} from '@/history/format';

const PAGE_SIZE = 50;

interface JobDetail { job: Job; messages: JobMessage[] }

function Stat({ label, value, tone }: { label: string; value: number; tone?: 'success' | 'destructive' | 'warning' | 'muted' }) {
  return (
    <div className="min-w-[88px]">
      <div className="text-[11px] font-semibold uppercase tracking-[0.12em] text-muted-foreground">{label}</div>
      <div className={cn(
        'mt-0.5 font-serif text-[24px] font-semibold tabular-nums leading-8',
        tone === 'success' && 'text-success', tone === 'destructive' && 'text-destructive',
        tone === 'warning' && 'text-warning', tone === 'muted' && 'text-muted-foreground',
      )}>{value.toLocaleString()}</div>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="min-w-0">
      <dt className="text-[11.5px] font-medium text-muted-foreground">{label}</dt>
      <dd className="mt-0.5 truncate text-[13.5px]">{children}</dd>
    </div>
  );
}

function Summary({ job }: { job: Job }) {
  const pending = pendingCount(job);
  return (
    <section aria-label="Summary" className="rounded-lg border bg-card">
      <div className="flex flex-wrap gap-x-8 gap-y-4 border-b px-5 py-4">
        <Stat label="Sent" value={job.sentCount} tone="success" />
        <Stat label="Failed" value={job.failedCount} tone={job.failedCount ? 'destructive' : 'muted'} />
        <Stat label="Skipped" value={job.skippedCount} tone={job.skippedCount ? 'warning' : 'muted'} />
        <Stat label="Pending" value={pending} tone="muted" />
        <div className="ml-auto self-end text-right text-[12.5px] text-muted-foreground">
          {job.totalMessages.toLocaleString()} message{job.totalMessages === 1 ? '' : 's'} to{' '}
          {job.totalRecipients.toLocaleString()} recipient{job.totalRecipients === 1 ? '' : 's'}
          <CountsBar job={job} className="mt-2 w-56" />
        </div>
      </div>
      <dl className="grid gap-x-8 gap-y-3 px-5 py-4 sm:grid-cols-2 lg:grid-cols-4">
        <Field label="From">{job.senderEmail}</Field>
        <Field label="Mode"><ModeBadge mode={job.mode} /></Field>
        <Field label="Status"><JobStatusBadge status={job.status} /></Field>
        <Field label="Created by">{job.userName ?? '—'}</Field>
        <Field label="Created"><span title={absoluteTime(job.createdAt)}>{absoluteTime(job.createdAt)}</span></Field>
        <Field label="Started">{absoluteTime(job.startedAt)}</Field>
        <Field label="Finished">{absoluteTime(job.finishedAt)}</Field>
        <Field label="Job ID"><span className="font-mono text-[12px]">{job.id}</span></Field>
      </dl>
    </section>
  );
}

const FILTERS: { id: MessageFilter; label: string }[] = [
  { id: 'all', label: 'All' },
  { id: 'sent', label: 'Sent' },
  { id: 'failed', label: 'Failed' },
  { id: 'skipped', label: 'Skipped' },
  { id: 'pending', label: 'Pending' },
];

function CopyList({ label, list }: { label: string; list: string | null }) {
  const emails = parseEmailList(list);
  if (!emails.length) return null;
  return (
    <div className="text-[12.5px]">
      <span className="font-medium text-muted-foreground">{label} ({emails.length}): </span>
      <span className="break-all">{emails.join(', ')}</span>
    </div>
  );
}

function MessagesTable({ messages }: { messages: JobMessage[] }) {
  const [filter, setFilter] = useState<MessageFilter>('all');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(0);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());

  const counts = useMemo(() => Object.fromEntries(FILTERS.map(f => [f.id, messages.filter(m => matchesFilter(m, f.id)).length])) as Record<MessageFilter, number>, [messages]);
  const rows = useMemo(
    () => messages.filter(m => matchesFilter(m, filter) && matchesSearch(m, search)).sort((a, b) => a.seq - b.seq),
    [messages, filter, search],
  );
  const pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
  useEffect(() => { setPage(0); }, [filter, search]);
  const current = Math.min(page, pages - 1);
  const visible = rows.slice(current * PAGE_SIZE, current * PAGE_SIZE + PAGE_SIZE);

  const toggle = (id: string) => setExpanded(prev => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  return (
    <section aria-labelledby="msgs-h" className="space-y-3">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <h2 id="msgs-h" className="text-[19px] font-semibold">Recipients</h2>
        <div className="relative w-full max-w-xs">
          <Search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" aria-hidden />
          <Input type="search" value={search} onChange={e => setSearch(e.target.value)} placeholder="Search address or error" aria-label="Search recipients" className="pl-8" />
        </div>
      </div>

      <div role="radiogroup" aria-label="Filter by status" className="flex flex-wrap gap-1.5">
        {FILTERS.map(f => (
          <button
            key={f.id}
            type="button"
            role="radio"
            aria-checked={filter === f.id}
            onClick={() => setFilter(f.id)}
            className={cn(
              'inline-flex h-8 items-center gap-1.5 rounded-full border px-3 text-[12.5px] font-medium transition-colors',
              'focus-visible:outline-2 focus-visible:outline-offset-1 focus-visible:outline-ring',
              filter === f.id ? 'border-primary bg-primary text-primary-foreground' : 'bg-card text-muted-foreground hover:text-foreground',
            )}
          >
            {f.label}
            <span className={cn('tabular-nums', filter === f.id ? 'opacity-80' : 'text-muted-foreground')}>{counts[f.id]}</span>
          </button>
        ))}
      </div>

      <div className="overflow-hidden rounded-lg border bg-card">
        <Table>
          <TableHeader className="bg-muted/50">
            <TableRow className="hover:bg-transparent">
              <TableHead className="w-14 pl-4">#</TableHead>
              <TableHead className="min-w-[220px]">To</TableHead>
              <TableHead>CC / BCC</TableHead>
              <TableHead>Status</TableHead>
              <TableHead className="text-right">Attempts</TableHead>
              <TableHead className="min-w-[200px]">Error</TableHead>
              <TableHead>Sent</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {visible.map(m => {
              const cc = countAddresses(m.cc);
              const bcc = countAddresses(m.bcc);
              const canExpand = cc + bcc > 0;
              const isOpen = expanded.has(m.id);
              return (
                <Fragment key={m.id}>
                  <TableRow className={cn(isOpen && 'border-b-0')}>
                    <TableCell className="pl-4 tabular-nums text-muted-foreground">{m.seq}</TableCell>
                    <TableCell className="max-w-[320px]">
                      <span className="block truncate text-[13px]" title={m.to}>{m.to}</span>
                      {m.recipientCount > 1 && <span className="text-[11.5px] text-muted-foreground">{m.recipientCount} recipients in this message</span>}
                    </TableCell>
                    <TableCell>
                      {canExpand ? (
                        <button
                          type="button"
                          onClick={() => toggle(m.id)}
                          aria-expanded={isOpen}
                          className="inline-flex items-center gap-1 rounded px-1 text-[12.5px] tabular-nums text-foreground hover:bg-muted focus-visible:outline-2 focus-visible:outline-ring"
                        >
                          {cc} / {bcc}
                          <ChevronDown className={cn('size-3.5 text-muted-foreground transition-transform', isOpen && 'rotate-180')} aria-hidden />
                          <span className="sr-only">{isOpen ? 'Hide' : 'Show'} copy addresses</span>
                        </button>
                      ) : <span className="text-[12.5px] text-muted-foreground">—</span>}
                    </TableCell>
                    <TableCell><MessageStatusBadge status={m.status} /></TableCell>
                    <TableCell className="text-right tabular-nums">{m.attempts}</TableCell>
                    <TableCell className="max-w-[360px] whitespace-normal text-[12.5px] text-destructive">
                      {m.error ? (
                        <>
                          {m.error}
                          {m.retryable === false && <span className="ml-1 text-muted-foreground">(permanent)</span>}
                        </>
                      ) : <span className="text-muted-foreground">—</span>}
                    </TableCell>
                    <TableCell className="text-[12.5px] text-muted-foreground"><TimeAgo iso={m.sentAt} /></TableCell>
                  </TableRow>
                  {isOpen && (
                    <TableRow className="bg-muted/30 hover:bg-muted/30">
                      <TableCell />
                      <TableCell colSpan={6} className="space-y-1 whitespace-normal pb-3 pt-0">
                        <CopyList label="CC" list={m.cc} />
                        <CopyList label="BCC" list={m.bcc} />
                      </TableCell>
                    </TableRow>
                  )}
                </Fragment>
              );
            })}
            {visible.length === 0 && (
              <TableRow className="hover:bg-transparent">
                <TableCell colSpan={7} className="py-10 text-center text-[13px] text-muted-foreground">
                  <SearchX className="mx-auto mb-2 size-5" aria-hidden />
                  No recipients match this filter.
                </TableCell>
              </TableRow>
            )}
          </TableBody>
        </Table>
      </div>

      {rows.length > PAGE_SIZE && (
        <nav aria-label="Recipient pages" className="flex items-center justify-between gap-3 text-[12.5px] text-muted-foreground">
          <span className="tabular-nums">
            {current * PAGE_SIZE + 1}–{Math.min(rows.length, (current + 1) * PAGE_SIZE)} of {rows.length}
          </span>
          <div className="flex items-center gap-1">
            <Button variant="outline" size="sm" onClick={() => setPage(current - 1)} disabled={current === 0}><ChevronLeft /> Previous</Button>
            <span className="px-2 tabular-nums">Page {current + 1} of {pages}</span>
            <Button variant="outline" size="sm" onClick={() => setPage(current + 1)} disabled={current >= pages - 1}>Next <ChevronRight /></Button>
          </div>
        </nav>
      )}
    </section>
  );
}

function DetailSkeleton() {
  return (
    <>
      <PageHeader title="Send details" />
      <PageBody className="space-y-6">
        <Skeleton className="h-40 w-full" />
        <Skeleton className="h-8 w-72" />
        <Skeleton className="h-64 w-full" />
      </PageBody>
    </>
  );
}

export default function JobDetailPage() {
  const { jobId = '' } = useParams();
  const navigate = useNavigate();
  const qc = useQueryClient();

  const query = useQuery({
    queryKey: ['jobs', 'detail', jobId],
    queryFn: ({ signal }) => api.get<JobDetail>(`/jobs/${encodeURIComponent(jobId)}`, { signal }),
    enabled: !!jobId,
    // Poll only while the job can change by itself (queued/sending); a paused job waits for someone to act.
    refetchInterval: q => (q.state.data && shouldPoll(q.state.data.job.status) ? 5_000 : false),
  });

  const invalidate = () => qc.invalidateQueries({ queryKey: ['jobs'] });

  const retry = useMutation({
    mutationFn: (_failed: number) => api.post<{ job: Job }>(`/jobs/${encodeURIComponent(jobId)}/retry-failed`),
    onSuccess: ({ job }, failed) => {
      toast.success(`Re-queued ${failed} failed message${failed === 1 ? '' : 's'}. Resuming the send.`);
      void invalidate();
      navigate(`/send/${job?.id || jobId}`);
    },
    onError: err => toast.error(`Could not retry: ${errorMessage(err)}`),
  });

  const cancel = useMutation({
    mutationFn: () => api.post<{ job: Job }>(`/jobs/${encodeURIComponent(jobId)}/cancel`),
    onSuccess: ({ job }) => {
      qc.setQueryData<JobDetail>(['jobs', 'detail', jobId], prev => (prev ? { ...prev, job } : prev));
      toast.success('Send cancelled. Messages not yet sent will not go out.');
      void invalidate();
    },
    onError: err => toast.error(`Could not cancel: ${errorMessage(err)}`),
  });

  if (query.isLoading) return <DetailSkeleton />;

  if (query.isError || !query.data) {
    const notFound = query.error instanceof ApiClientError && (query.error.status === 404 || query.error.status === 403);
    return (
      <>
        <PageHeader title="Send details" actions={<Button variant="outline" asChild><Link to="/history"><ArrowLeft /> Sent history</Link></Button>} />
        <PageBody>
          <EmptyState
            icon={<SearchX className="size-5" />}
            title={notFound ? 'This send was not found' : 'Could not load this send'}
            action={notFound ? <Button asChild><Link to="/history">Back to Sent history</Link></Button> : <Button variant="outline" onClick={() => query.refetch()}><RotateCcw /> Retry</Button>}
          >
            {notFound ? 'It may have been removed, or it belongs to another user.' : errorMessage(query.error)}
          </EmptyState>
        </PageBody>
      </>
    );
  }

  const { job, messages } = query.data;
  const active = isActive(job.status);
  const canRetry = job.failedCount > 0 && job.status !== 'sending' && job.status !== 'queued';

  return (
    <>
      <PageHeader
        title={job.subject || '(no subject)'}
        description={
          <span className="flex flex-wrap items-center gap-x-2 gap-y-1">
            <Link to="/history" className="inline-flex items-center gap-1 text-foreground underline-offset-2 hover:underline">
              <ArrowLeft className="size-3.5" aria-hidden /> Sent history
            </Link>
            <span aria-hidden>·</span>
            <span>From {job.senderEmail}</span>
            <span aria-hidden>·</span>
            <span>Created <TimeAgo iso={job.createdAt} /></span>
          </span>
        }
        actions={
          <>
            {active && (
              <Button asChild>
                <Link to={`/send/${job.id}`}><RadioTower /> Open live view</Link>
              </Button>
            )}
            {canRetry && (
              <Button variant="outline" onClick={() => retry.mutate(job.failedCount)} disabled={retry.isPending}>
                {retry.isPending ? <Loader2 className="animate-spin" /> : <RotateCcw />} Retry {job.failedCount} failed
              </Button>
            )}
            <Button variant="outline" asChild>
              <a href={`/api/jobs/${encodeURIComponent(job.id)}/export.csv`} download>
                <Download /> Export CSV
              </a>
            </Button>
            {active && (
              <AlertDialog>
                <AlertDialogTrigger asChild>
                  <Button variant="outline" className="text-destructive hover:text-destructive" disabled={cancel.isPending}>
                    {cancel.isPending ? <Loader2 className="animate-spin" /> : <Ban />} Cancel send
                  </Button>
                </AlertDialogTrigger>
                <AlertDialogContent>
                  <AlertDialogHeader>
                    <AlertDialogTitle>Cancel this send?</AlertDialogTitle>
                    <AlertDialogDescription>
                      {pendingCount(job).toLocaleString()} message{pendingCount(job) === 1 ? '' : 's'} not yet sent from {job.senderEmail} will not go out.
                      Messages already delivered cannot be recalled. This cannot be undone.
                    </AlertDialogDescription>
                  </AlertDialogHeader>
                  <AlertDialogFooter>
                    <AlertDialogCancel>Keep sending</AlertDialogCancel>
                    <AlertDialogAction variant="destructive" onClick={() => cancel.mutate()}>Cancel send</AlertDialogAction>
                  </AlertDialogFooter>
                </AlertDialogContent>
              </AlertDialog>
            )}
          </>
        }
      />
      <PageBody className="space-y-8">
        <Summary job={job} />
        <MessagesTable messages={messages} />
      </PageBody>
    </>
  );
}
