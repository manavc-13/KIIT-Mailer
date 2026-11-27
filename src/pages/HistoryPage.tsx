// Sent history: the signed-in user's send jobs (admins can switch to all users).
import { useMemo, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { useInfiniteQuery } from '@tanstack/react-query';
import { History, Loader2, PenLine, RotateCcw, Search, SearchX } from 'lucide-react';
import { PageBody, PageHeader } from '@/components/AppShell';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Skeleton } from '@/components/ui/skeleton';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { api, errorMessage } from '@/lib/api';
import { useAuth } from '@/lib/auth';
import { cn } from '@/lib/utils';
import type { Job } from '../../shared/types';
import { CountsBar, CountsInline, EmptyState, JobStatusBadge, ModeBadge, TimeAgo } from '@/history/components';
import { shouldPoll } from '@/history/format';

type Scope = 'mine' | 'all';
const SCOPE_KEY = 'kiit-mailer.history.scope';

function readScope(): Scope {
  try { return window.localStorage.getItem(SCOPE_KEY) === 'all' ? 'all' : 'mine'; } catch { return 'mine'; }
}

function ScopeToggle({ value, onChange }: { value: Scope; onChange: (s: Scope) => void }) {
  return (
    <div role="radiogroup" aria-label="Whose sends to show" className="inline-flex rounded-md border bg-muted p-0.5">
      {([['mine', 'Only mine'], ['all', 'All users']] as const).map(([v, label]) => (
        <button
          key={v}
          type="button"
          role="radio"
          aria-checked={value === v}
          onClick={() => onChange(v)}
          className={cn(
            'h-8 rounded-[5px] px-3 text-[13px] font-medium text-muted-foreground transition-colors hover:text-foreground',
            'focus-visible:outline-2 focus-visible:outline-offset-1 focus-visible:outline-ring',
            value === v && 'bg-card text-foreground shadow-[0_1px_2px_rgba(12,26,40,0.12)]',
          )}
        >
          {label}
        </button>
      ))}
    </div>
  );
}

function LoadingRows({ cols }: { cols: number }) {
  return (
    <>
      {Array.from({ length: 6 }, (_, i) => (
        <TableRow key={i} className="hover:bg-transparent">
          {Array.from({ length: cols }, (_, j) => (
            <TableCell key={j} className="py-4"><Skeleton className={cn('h-4', j === 0 ? 'w-56' : 'w-20')} /></TableCell>
          ))}
        </TableRow>
      ))}
    </>
  );
}

export default function HistoryPage() {
  const { isAdmin } = useAuth();
  const navigate = useNavigate();
  const [scopePref, setScopePref] = useState<Scope>(readScope);
  const scope: Scope = isAdmin ? scopePref : 'mine';
  const [search, setSearch] = useState('');

  const setScope = (s: Scope) => {
    setScopePref(s);
    try { window.localStorage.setItem(SCOPE_KEY, s); } catch { /* ignore */ }
  };

  const query = useInfiniteQuery({
    queryKey: ['jobs', 'list', scope],
    queryFn: ({ pageParam, signal }) => {
      const qs = new URLSearchParams({ scope });
      if (pageParam) qs.set('cursor', pageParam);
      return api.get<{ jobs: Job[]; nextCursor: string | null }>(`/jobs?${qs}`, { signal });
    },
    initialPageParam: null as string | null,
    getNextPageParam: last => last.nextCursor,
    // Keep progress fresh while any listed job is still running.
    refetchInterval: q => (q.state.data?.pages.some(p => p.jobs.some(j => shouldPoll(j.status))) ? 10_000 : false),
  });

  const jobs = useMemo(() => {
    const seen = new Set<string>();
    return (query.data?.pages ?? []).flatMap(p => p.jobs).filter(j => (seen.has(j.id) ? false : (seen.add(j.id), true)));
  }, [query.data]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return jobs;
    return jobs.filter(j => j.subject.toLowerCase().includes(q) || j.senderEmail.toLowerCase().includes(q) || (j.userName ?? '').toLowerCase().includes(q));
  }, [jobs, search]);

  const showUser = isAdmin && scope === 'all';
  const cols = showUser ? 7 : 6;
  const open = (j: Job) => navigate(`/history/${j.id}`);

  const empty = !query.isLoading && !query.isError && jobs.length === 0;

  return (
    <>
      <PageHeader
        title="Sent history"
        description={showUser ? 'Every send job across all users, newest first.' : 'Your send jobs, newest first. Open one to see each recipient, retry failures or export a report.'}
        actions={isAdmin ? <ScopeToggle value={scope} onChange={setScope} /> : undefined}
      />
      <PageBody>
        {empty ? (
          <EmptyState
            icon={<History className="size-5" />}
            title={scope === 'all' ? 'No mail has been sent yet' : 'You have not sent any mail yet'}
            action={<Button asChild><Link to="/compose"><PenLine /> Compose a message</Link></Button>}
          >
            Each send appears here with per-recipient delivery status, so you can retry failures and export a report.
          </EmptyState>
        ) : (
          <div className="space-y-4">
            <div className="flex flex-wrap items-center gap-3">
              <div className="relative w-full max-w-sm">
                <Search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" aria-hidden />
                <Input
                  type="search"
                  value={search}
                  onChange={e => setSearch(e.target.value)}
                  placeholder={showUser ? 'Search subject, sender or person' : 'Search subject or sender'}
                  aria-label={showUser ? 'Search loaded sends by subject, sender or person' : 'Search loaded sends by subject or sender'}
                  className="pl-8"
                />
              </div>
              {jobs.length > 0 && (
                <span className="text-[12.5px] text-muted-foreground">
                  {search ? `${filtered.length} of ${jobs.length} loaded` : `${jobs.length} loaded`}
                  {query.hasNextPage ? ', more available' : ''}
                </span>
              )}
            </div>

            {query.isError ? (
              <div role="alert" className="flex items-center gap-3 rounded-md border border-destructive/30 bg-destructive/5 px-4 py-3 text-[13px] text-destructive">
                <span className="flex-1">Could not load your history. {errorMessage(query.error)}</span>
                <Button variant="outline" size="sm" onClick={() => query.refetch()}><RotateCcw /> Retry</Button>
              </div>
            ) : (
              <div className="overflow-hidden rounded-lg border bg-card">
                <Table>
                  <TableHeader className="bg-muted/50">
                    <TableRow className="hover:bg-transparent">
                      <TableHead className="min-w-[240px] pl-4">Subject</TableHead>
                      <TableHead className="hidden lg:table-cell">From</TableHead>
                      <TableHead className="hidden md:table-cell">Mode</TableHead>
                      <TableHead>Status</TableHead>
                      <TableHead className="min-w-[170px]">Delivery</TableHead>
                      <TableHead>Created</TableHead>
                      {showUser && <TableHead className="hidden md:table-cell">Sent by</TableHead>}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {query.isLoading && <LoadingRows cols={cols} />}
                    {filtered.map(j => (
                      <TableRow
                        key={j.id}
                        onClick={e => { if (!(e.target as HTMLElement).closest('a,button')) open(j); }}
                        className="cursor-pointer"
                      >
                        <TableCell className="max-w-[360px] py-3 pl-4">
                          <Link
                            to={`/history/${j.id}`}
                            className="block truncate font-medium text-foreground underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring"
                          >
                            {j.subject || '(no subject)'}
                          </Link>
                          <span className="block truncate text-[12px] text-muted-foreground lg:hidden">{j.senderEmail}</span>
                        </TableCell>
                        <TableCell className="hidden max-w-[220px] truncate text-[13px] text-muted-foreground lg:table-cell">{j.senderEmail}</TableCell>
                        <TableCell className="hidden md:table-cell"><ModeBadge mode={j.mode} /></TableCell>
                        <TableCell><JobStatusBadge status={j.status} /></TableCell>
                        <TableCell>
                          <CountsInline job={j} />
                          <CountsBar job={j} className="mt-1.5 max-w-[180px]" />
                        </TableCell>
                        <TableCell className="text-[13px] text-muted-foreground"><TimeAgo iso={j.createdAt} /></TableCell>
                        {showUser && <TableCell className="hidden text-[13px] md:table-cell">{j.userName ?? '—'}</TableCell>}
                      </TableRow>
                    ))}
                    {!query.isLoading && filtered.length === 0 && jobs.length > 0 && (
                      <TableRow className="hover:bg-transparent">
                        <TableCell colSpan={cols} className="py-10 text-center text-[13px] text-muted-foreground">
                          <SearchX className="mx-auto mb-2 size-5" aria-hidden />
                          No loaded sends match “{search}”.{query.hasNextPage ? ' Load more to search older sends.' : ''}
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </div>
            )}

            {query.hasNextPage && (
              <div className="flex justify-center">
                <Button variant="outline" onClick={() => query.fetchNextPage()} disabled={query.isFetchingNextPage}>
                  {query.isFetchingNextPage ? <><Loader2 className="animate-spin" /> Loading…</> : 'Load more'}
                </Button>
              </div>
            )}
          </div>
        )}
      </PageBody>
    </>
  );
}
