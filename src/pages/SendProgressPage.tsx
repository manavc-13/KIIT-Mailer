// Live send progress for one job (/send/:jobId). While the job is queued or sending and this page
// is open, it pumps POST /api/jobs/:id/run (each call sends for up to ~30-35 s), with backoff on
// network errors and a short wait when a call made no progress. Pause / resume / cancel /
// retry-failed are server state changes, so the job can be left and picked up again from Sent
// history or the "Sending in progress" banner.
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Link, useParams } from 'react-router-dom';
import { useQuery, useQueryClient } from '@tanstack/react-query';
import { toast } from 'sonner';
import {
  Pause, Play, XCircle, RotateCcw, Download, Loader2, CheckCircle2, AlertTriangle, History, PenLine, Info, KeyRound, Gauge,
} from 'lucide-react';
import { PageBody, PageHeader } from '@/components/AppShell';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter,
  AlertDialogHeader, AlertDialogTitle,
} from '@/components/ui/alert-dialog';
import { cn } from '@/lib/utils';
import { api, ApiClientError, errorMessage, sendErrorTitle } from '@/lib/api';
import type { Job, JobMessage, MessageStatus, RunJobResponse } from '../../shared/types';
import { Segmented, StatPill, Chip, type Tone } from '@/compose/ui';

type JobData = { job: Job; messages: JobMessage[] };
type Filter = 'all' | 'sent' | 'failed' | 'pending' | 'skipped';
type PauseReason = 'user' | 'quota' | 'auth' | 'connection' | 'error' | null;

const ACTIVE: Job['status'][] = ['queued', 'sending'];
const AUTH_RE = /app password|invalid login|535|5\.7\.8|authentication|username and password|credentials/i;
const QUOTA_RE = /quota|daily (user )?sending limit|5\.4\.5|limit exceeded|too many/i;

/** Wait before the next /run call when one returned nothing processed and not done (e.g. another tab holds the claim). */
const IDLE_RUN_DELAY_MS = 5000;
/** The pump merges each /run result into the cache, so the full job detail only needs an occasional refresh. */
const DETAIL_REFRESH_MS = 15_000;

const sleep = (ms: number, signal: AbortSignal) => new Promise<void>(resolve => {
  const t = setTimeout(resolve, ms);
  signal.addEventListener('abort', () => { clearTimeout(t); resolve(); }, { once: true });
});

function formatElapsed(ms: number): string {
  const s = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), sec = s % 60;
  return h ? `${h}h ${String(m).padStart(2, '0')}m` : `${m}:${String(sec).padStart(2, '0')}`;
}

export default function SendProgressPage() {
  const { jobId = '' } = useParams();
  const qc = useQueryClient();
  const key = useMemo(() => ['jobs', 'detail', jobId], [jobId]);
  const [pumping, setPumping] = useState(false);
  const [pauseReason, setPauseReason] = useState<PauseReason>(null);
  const [runError, setRunError] = useState<string | null>(null);
  /** Plain-English reason from the server for an automatic pause (RunJobResponse.error). */
  const [pauseDetail, setPauseDetail] = useState<string | null>(null);
  const [busyAction, setBusyAction] = useState<null | 'pause' | 'cancel' | 'retry'>(null);
  const [confirmCancel, setConfirmCancel] = useState(false);

  const q = useQuery({
    queryKey: key,
    queryFn: () => api.get<JobData>(`/jobs/${encodeURIComponent(jobId)}`),
    enabled: !!jobId,
    // Occasional full refresh while pumping; each /run response is merged into this cache directly.
    refetchInterval: pumping ? DETAIL_REFRESH_MS : false,
  });
  const job = q.data?.job;
  const messages = q.data?.messages ?? [];

  const mergeRun = useCallback((res: RunJobResponse) => {
    qc.setQueryData<JobData>(key, old => {
      if (!old) return { job: res.job, messages: res.processed };
      const byId = new Map(res.processed.map(m => [m.id, m]));
      return { job: res.job, messages: old.messages.map(m => byId.get(m.id) ?? m) };
    });
  }, [qc, key]);

  const setJob = useCallback((j: Job) => {
    qc.setQueryData<JobData>(key, old => (old ? { ...old, job: j } : old));
  }, [qc, key]);

  // ---- the pump loop
  const loopRef = useRef<{ ctrl: AbortController } | null>(null);

  const stopLoop = useCallback(() => {
    loopRef.current?.ctrl.abort();
    loopRef.current = null;
    setPumping(false);
  }, []);

  const startLoop = useCallback(() => {
    if (loopRef.current || !jobId) return;
    const me = { ctrl: new AbortController() };
    loopRef.current = me;
    setPumping(true);
    setPauseReason(null);
    setPauseDetail(null);
    setRunError(null);
    (async () => {
      let failures = 0;
      try {
        while (loopRef.current === me) {
          try {
            const res = await api.post<RunJobResponse>(`/jobs/${encodeURIComponent(jobId)}/run`, {}, { signal: me.ctrl.signal });
            if (loopRef.current !== me) break;
            failures = 0;
            setRunError(null);
            mergeRun(res);
            // done = stop looping (completed, cancelled, or an automatic pause with `error`).
            if (res.done || !ACTIVE.includes(res.job.status)) {
              if (res.job.status === 'paused') {
                const why = res.error || res.processed.map(m => m.error || '').join(' ');
                if (res.error) setPauseDetail(res.error);
                setPauseReason(prev => prev ?? (res.quotaExhausted || (!AUTH_RE.test(why) && QUOTA_RE.test(why)) ? 'quota'
                  : AUTH_RE.test(why) ? 'auth' : res.error ? 'error' : null));
              }
              break;
            }
            // Not done but nothing processed (e.g. another tab is sending this job, or everything left is
            // waiting for a retry): wait before asking again instead of flooding /run.
            if (res.processed.length === 0) await sleep(IDLE_RUN_DELAY_MS, me.ctrl.signal);
          } catch (e) {
            if (me.ctrl.signal.aborted || loopRef.current !== me) break;
            if (e instanceof ApiClientError && e.status >= 400 && e.status < 500 && e.status !== 408 && e.status !== 429) {
              // /run reports quota and sign-in problems in its 200 body (handled above), so any 4xx here is
              // a plain stop: the server's message is shown as-is.
              setRunError(e.message);
              setPauseReason('error');
              break;
            }
            failures++;
            if (failures > 8) {
              setRunError('Lost contact with the server. Check your connection, then press Resume.');
              setPauseReason('connection');
              break;
            }
            const delay = Math.min(30_000, 2000 * 2 ** (failures - 1));
            setRunError(`Connection problem (${errorMessage(e)}). Retrying in ${Math.round(delay / 1000)} s…`);
            await sleep(delay, me.ctrl.signal);
          }
        }
      } finally {
        if (loopRef.current === me) { loopRef.current = null; setPumping(false); }
        void qc.invalidateQueries({ queryKey: key });
        void qc.invalidateQueries({ queryKey: ['jobs', 'active'] });
      }
    })();
  }, [jobId, mergeRun, qc, key]);

  useEffect(() => () => stopLoop(), [stopLoop]);

  // Auto-start whenever the job is queued/sending and nothing is pumping it from this tab.
  // Deferred with a cleared timeout so React StrictMode's mount → unmount → mount runs only one pump.
  const status = job?.status;
  useEffect(() => {
    if (!(status && ACTIVE.includes(status) && !pauseReason)) return;
    const t = setTimeout(startLoop, 0);
    return () => clearTimeout(t);
  }, [status, pauseReason, startLoop]);

  // Infer why a job is paused when we arrive at an already-paused job.
  useEffect(() => {
    if (status !== 'paused' || pauseReason) return;
    const errs = messages.filter(m => m.error).slice(-5).map(m => m.error).join(' ');
    if (AUTH_RE.test(errs)) setPauseReason('auth');
    else if (QUOTA_RE.test(errs)) setPauseReason('quota');
  }, [status, messages, pauseReason]);

  // Warn before leaving while this tab is the one sending.
  useEffect(() => {
    if (!pumping) return;
    const h = (e: BeforeUnloadEvent) => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', h);
    return () => window.removeEventListener('beforeunload', h);
  }, [pumping]);

  // ---- actions
  const act = async (which: 'pause' | 'cancel' | 'retry') => {
    setBusyAction(which);
    try {
      const path = which === 'retry' ? 'retry-failed' : which;
      const res = await api.post<{ job: Job }>(`/jobs/${encodeURIComponent(jobId)}/${path}`);
      if (which !== 'retry') stopLoop();
      setJob(res.job);
      if (which === 'pause') { setPauseReason('user'); toast('Sending paused'); }
      if (which === 'cancel') toast.success('Send cancelled. Messages not yet sent will not go out.');
      if (which === 'retry') { setPauseReason(null); toast.success('Failed messages re-queued'); startLoop(); }
      void qc.invalidateQueries({ queryKey: key });
      void qc.invalidateQueries({ queryKey: ['jobs', 'active'] });
    } catch (e) {
      toast.error(sendErrorTitle(e, `Could not ${which === 'retry' ? 'retry failed messages' : which}`), { description: errorMessage(e) });
    } finally {
      setBusyAction(null);
    }
  };

  // Resume = call /run again; the server sets the job back to sending.
  const resume = () => { setPauseReason(null); setPauseDetail(null); startLoop(); };

  // ---- render
  if (q.isLoading) {
    return (
      <>
        <PageHeader title="Sending" />
        <PageBody className="space-y-4"><Skeleton className="h-28 w-full" /><Skeleton className="h-64 w-full" /></PageBody>
      </>
    );
  }
  if (q.error || !job) {
    return (
      <>
        <PageHeader title="Sending" />
        <PageBody>
          <div role="alert" className="max-w-xl rounded-lg border bg-card p-6 shadow-xs">
            <h2 className="text-lg font-semibold">This job could not be opened</h2>
            <p className="mt-1 text-sm text-muted-foreground">{q.error ? errorMessage(q.error) : 'It may have been removed, or it belongs to another user.'}</p>
            <Button asChild variant="outline" className="mt-4"><Link to="/history"><History /> Sent history</Link></Button>
          </div>
        </PageBody>
      </>
    );
  }

  const total = job.totalMessages || messages.length;
  const doneCount = job.sentCount + job.failedCount + job.skippedCount;
  const pending = Math.max(0, total - doneCount);
  const pct = total ? Math.round((doneCount / total) * 100) : 0;
  const finished = job.status === 'completed' || job.status === 'cancelled';
  const paused = !pumping && !finished && (job.status === 'paused' || !!pauseReason);

  return (
    <>
      <PageHeader
        title={finished ? (job.status === 'completed' ? 'Sending complete' : 'Send cancelled') : paused ? 'Sending paused' : 'Sending…'}
        description={<>“{job.subject}” from <strong className="font-medium text-foreground">{job.senderEmail}</strong> · {job.mode === 'personal' ? 'personalized' : 'BCC broadcast'}</>}
        actions={
          <Button asChild variant="outline" size="sm">
            <a href={`/api/jobs/${encodeURIComponent(job.id)}/export.csv`} download><Download /> Export CSV</a>
          </Button>
        }
      />
      <PageBody className="space-y-5">
        {!finished && (
          <p className="flex items-start gap-2 rounded-md border bg-gold-soft px-4 py-2.5 text-sm text-accent-foreground">
            <Info className="mt-0.5 size-4 shrink-0" />
            Keep this tab open while sending. You can leave and resume later from Sent history; nothing is lost.
          </p>
        )}

        <PauseNotice reason={paused ? pauseReason : null} detail={pauseDetail} runError={runError} pumping={pumping} />

        <section aria-label="Progress" className="rounded-lg border bg-card p-5 shadow-xs">
          <div className="flex flex-wrap items-end justify-between gap-3">
            <div>
              <div className="font-serif text-3xl font-semibold tabular-nums">{pct}%</div>
              <div className="text-sm text-muted-foreground">
                {doneCount.toLocaleString()} of {total.toLocaleString()} messages processed · {job.totalRecipients.toLocaleString()} recipients
              </div>
            </div>
            <Elapsed job={job} running={pumping} />
          </div>
          <div
            role="progressbar" aria-label="Send progress" aria-valuemin={0} aria-valuemax={100} aria-valuenow={pct}
            className="mt-4 flex h-3 w-full overflow-hidden rounded-full bg-muted ring-1 ring-inset ring-border"
          >
            <div className="h-full bg-success transition-[width]" style={{ width: `${total ? (job.sentCount / total) * 100 : 0}%` }} />
            <div className="h-full bg-destructive transition-[width]" style={{ width: `${total ? (job.failedCount / total) * 100 : 0}%` }} />
            <div className="h-full bg-warning transition-[width]" style={{ width: `${total ? (job.skippedCount / total) * 100 : 0}%` }} />
          </div>
          <div className="mt-4 flex flex-wrap gap-2">
            <StatPill label="Sent" value={job.sentCount.toLocaleString()} tone="ok" />
            <StatPill label="Failed" value={job.failedCount.toLocaleString()} tone={job.failedCount ? 'danger' : 'neutral'} />
            <StatPill label="Skipped" value={job.skippedCount.toLocaleString()} tone={job.skippedCount ? 'warn' : 'neutral'} />
            <StatPill label="Pending" value={pending.toLocaleString()} tone="neutral" />
          </div>

          {!finished && (
            <div className="mt-5 flex flex-wrap items-center gap-2 border-t pt-4">
              {pumping || (ACTIVE.includes(job.status) && !pauseReason) ? (
                <Button type="button" variant="outline" onClick={() => act('pause')} disabled={!!busyAction}>
                  {busyAction === 'pause' ? <Loader2 className="animate-spin" /> : <Pause />} Pause
                </Button>
              ) : (
                <Button type="button" onClick={resume} disabled={!!busyAction}><Play /> Resume</Button>
              )}
              <Button type="button" variant="ghost" className="text-destructive hover:text-destructive" onClick={() => setConfirmCancel(true)} disabled={!!busyAction}>
                <XCircle /> Cancel send
              </Button>
              {/* Same rule as JobDetailPage: retrying is only offered once nothing is sending. */}
              {job.failedCount > 0 && !pumping && job.status !== 'sending' && (
                <Button type="button" variant="outline" onClick={() => act('retry')} disabled={!!busyAction} className="ml-auto">
                  {busyAction === 'retry' ? <Loader2 className="animate-spin" /> : <RotateCcw />} Retry {job.failedCount} failed
                </Button>
              )}
              {pumping && <span className="flex items-center gap-1.5 text-xs text-muted-foreground"><Loader2 className="size-3.5 animate-spin" /> Sending from this tab</span>}
            </div>
          )}
        </section>

        {finished && <CompletionCard job={job} onRetry={() => act('retry')} retrying={busyAction === 'retry'} />}

        <MessagesTable messages={messages} />
      </PageBody>

      <AlertDialog open={confirmCancel} onOpenChange={setConfirmCancel}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Cancel this send?</AlertDialogTitle>
            <AlertDialogDescription>
              {pending.toLocaleString()} message{pending === 1 ? '' : 's'} not yet sent will be dropped. Messages already sent cannot be recalled.
              To stop temporarily, use Pause instead.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Keep sending</AlertDialogCancel>
            <AlertDialogAction variant="destructive" onClick={() => act('cancel')}>Cancel send</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </>
  );
}

function Elapsed({ job, running }: { job: Job; running: boolean }) {
  const [now, setNow] = useState(() => Date.now());
  useEffect(() => {
    if (!running) return;
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, [running]);
  const start = Date.parse(job.startedAt ?? job.createdAt);
  const end = job.finishedAt ? Date.parse(job.finishedAt) : now;
  return (
    <div className="text-right">
      <div className="text-xs uppercase tracking-wide text-muted-foreground">Elapsed</div>
      <div className="font-mono text-lg tabular-nums">{Number.isFinite(start) ? formatElapsed(end - start) : '—'}</div>
    </div>
  );
}

function PauseNotice({ reason, detail, runError, pumping }: {
  reason: PauseReason; detail: string | null; runError: string | null; pumping: boolean;
}) {
  const said = detail ? <span className="mb-1 block text-foreground">{detail}</span> : null;
  if (reason === 'auth') {
    return (
      <Notice tone="danger" icon={KeyRound} title="Paused: the sender account's App Password was rejected">
        {said}
        Gmail refused to sign in, so sending stopped to avoid more failures. Next step: an administrator re-enters the App Password on the
        Sender accounts page, then press Resume here (or from Sent history).
      </Notice>
    );
  }
  if (reason === 'quota') {
    return (
      <Notice tone="warn" icon={Gauge} title="Paused: daily sending limit reached">
        {said}
        This account has used today's quota (Gmail counts To, CC and BCC). The remaining messages stay queued: resume tomorrow from Sent
        history, or send the rest from another account.
      </Notice>
    );
  }
  if (reason === 'error' && detail) {
    return (
      <Notice tone="danger" icon={AlertTriangle} title="Paused automatically">
        {said}
        Fix the cause (for a disabled sender, ask the administrator or compose again from another account; for a missing attachment,
        compose again with the file re-attached), then press Resume.
      </Notice>
    );
  }
  if (runError) {
    return (
      <Notice tone={pumping ? 'warn' : 'danger'} icon={AlertTriangle} title={pumping ? 'Having trouble reaching the server' : 'Sending stopped'}>
        {runError}
      </Notice>
    );
  }
  return null;
}

function Notice({ tone, icon: Icon, title, children }: { tone: 'warn' | 'danger'; icon: typeof Info; title: string; children: React.ReactNode }) {
  return (
    <div role="alert" className={cn('flex items-start gap-3 rounded-lg border px-4 py-3',
      tone === 'danger' ? 'border-destructive/30 bg-destructive/5' : 'border-warning/30 bg-warning-soft')}>
      <Icon className={cn('mt-0.5 size-5 shrink-0', tone === 'danger' ? 'text-destructive' : 'text-warning')} />
      <div>
        <div className="font-medium">{title}</div>
        <div className="mt-0.5 text-sm text-muted-foreground">{children}</div>
      </div>
    </div>
  );
}

function CompletionCard({ job, onRetry, retrying }: { job: Job; onRetry: () => void; retrying: boolean }) {
  const ok = job.status === 'completed' && job.failedCount === 0;
  return (
    <section aria-label="Summary" className="rounded-lg border bg-card p-5 shadow-xs">
      <div className="flex items-start gap-3">
        {ok ? <CheckCircle2 className="size-6 shrink-0 text-success" /> : <AlertTriangle className="size-6 shrink-0 text-warning" />}
        <div className="min-w-0 flex-1">
          <h2 className="text-lg font-semibold">
            {job.status === 'cancelled' ? 'Send cancelled' : ok ? 'All messages sent' : 'Finished with some failures'}
          </h2>
          <p className="mt-1 text-sm text-muted-foreground">
            {job.sentCount.toLocaleString()} sent · {job.failedCount.toLocaleString()} failed · {job.skippedCount.toLocaleString()} skipped
            {job.skippedCount > 0 && ' (unsubscribed, bounced or invalid addresses)'}.
          </p>
          <div className="mt-4 flex flex-wrap gap-2">
            <Button asChild variant="outline"><Link to={`/history/${job.id}`}><History /> View in history</Link></Button>
            <Button asChild><Link to="/compose"><PenLine /> Compose another</Link></Button>
            {job.failedCount > 0 && (
              <Button type="button" variant="outline" onClick={onRetry} disabled={retrying}>
                {retrying ? <Loader2 className="animate-spin" /> : <RotateCcw />} Retry failed
              </Button>
            )}
          </div>
        </div>
      </div>
    </section>
  );
}

const STATUS_CHIP: Record<MessageStatus, { label: string; tone: Tone }> = {
  queued: { label: 'Pending', tone: 'neutral' },
  sending: { label: 'Sending', tone: 'info' },
  sent: { label: 'Sent', tone: 'ok' },
  failed: { label: 'Failed', tone: 'danger' },
  skipped: { label: 'Skipped', tone: 'warn' },
};

const PAGE = 250;

function MessagesTable({ messages }: { messages: JobMessage[] }) {
  const [filter, setFilter] = useState<Filter>('all');
  const [limit, setLimit] = useState(PAGE);
  const counts = useMemo(() => ({
    all: messages.length,
    sent: messages.filter(m => m.status === 'sent').length,
    failed: messages.filter(m => m.status === 'failed').length,
    pending: messages.filter(m => m.status === 'queued' || m.status === 'sending').length,
    skipped: messages.filter(m => m.status === 'skipped').length,
  }), [messages]);
  const rows = useMemo(() => {
    const sorted = [...messages].sort((a, b) => a.seq - b.seq);
    if (filter === 'all') return sorted;
    if (filter === 'pending') return sorted.filter(m => m.status === 'queued' || m.status === 'sending');
    return sorted.filter(m => m.status === filter);
  }, [messages, filter]);

  return (
    <section aria-labelledby="messages-title" className="rounded-lg border bg-card shadow-xs">
      <header className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-3">
        <h2 id="messages-title" className="text-[17px] font-semibold">Recipients</h2>
        <Segmented<Filter>
          ariaLabel="Filter recipients" size="sm" value={filter} onChange={f => { setFilter(f); setLimit(PAGE); }}
          options={[
            { value: 'all', label: `All ${counts.all}` },
            { value: 'sent', label: `Sent ${counts.sent}` },
            { value: 'failed', label: `Failed ${counts.failed}` },
            { value: 'pending', label: `Pending ${counts.pending}` },
            ...(counts.skipped ? [{ value: 'skipped' as Filter, label: `Skipped ${counts.skipped}` }] : []),
          ]}
        />
      </header>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-muted/60 text-left text-xs text-muted-foreground">
            <tr>
              <th scope="col" className="w-14 px-4 py-2 font-medium">#</th>
              <th scope="col" className="px-4 py-2 font-medium">To</th>
              <th scope="col" className="w-28 px-4 py-2 font-medium">Status</th>
              <th scope="col" className="px-4 py-2 font-medium">Details</th>
            </tr>
          </thead>
          <tbody aria-live="polite" aria-relevant="text">
            {rows.slice(0, limit).map(m => {
              const chip = STATUS_CHIP[m.status];
              return (
                <tr key={m.id} className="border-t align-top">
                  <td className="px-4 py-2 tabular-nums text-muted-foreground">{m.seq}</td>
                  <td className="max-w-[22rem] px-4 py-2">
                    <div className="truncate font-medium" title={m.to}>{m.to}</div>
                    {(m.cc || m.bcc) && (
                      <div className="truncate text-xs text-muted-foreground" title={[m.cc && `CC ${m.cc}`, m.bcc && `BCC ${m.bcc}`].filter(Boolean).join(' · ')}>
                        {m.cc && `CC ${m.cc}`}{m.cc && m.bcc && ' · '}{m.bcc && `BCC ${m.bcc}`}
                      </div>
                    )}
                    {m.recipientCount > 1 && !m.cc && !m.bcc && <div className="text-xs text-muted-foreground">{m.recipientCount} recipients</div>}
                  </td>
                  <td className="px-4 py-2">
                    <Chip tone={chip.tone}>{m.status === 'sending' && <Loader2 className="size-3 animate-spin" />}{chip.label}</Chip>
                  </td>
                  <td className="px-4 py-2 text-xs">
                    {m.error ? <span className="text-destructive">{m.error}{m.retryable ? ' (will retry)' : ''}</span>
                      : m.sentAt ? <span className="text-muted-foreground">{new Date(m.sentAt).toLocaleTimeString()}{m.attempts > 1 ? ` · ${m.attempts} attempts` : ''}</span>
                        : null}
                  </td>
                </tr>
              );
            })}
            {rows.length === 0 && (
              <tr><td colSpan={4} className="px-4 py-8 text-center text-sm text-muted-foreground">Nothing here.</td></tr>
            )}
          </tbody>
        </table>
      </div>
      {rows.length > limit && (
        <div className="border-t px-5 py-3 text-center">
          <Button type="button" variant="ghost" size="sm" onClick={() => setLimit(l => l + PAGE * 4)}>
            Show more ({(rows.length - limit).toLocaleString()} remaining)
          </Button>
        </div>
      )}
    </section>
  );
}
