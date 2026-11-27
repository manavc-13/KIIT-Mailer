// Admin: audit log. Filter by action, user and date range; newest first with
// "Load more" (cursor); expandable details; CSV export of the current filter.
import { Fragment, useMemo, useState } from 'react';
import { useInfiniteQuery } from '@tanstack/react-query';
import { ChevronDown, ChevronRight, Download, Loader2, ScrollText, X } from 'lucide-react';
import { PageHeader, PageBody } from '@/components/AppShell';
import { api, errorMessage } from '@/lib/api';
import { cn } from '@/lib/utils';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import {
  Select, SelectContent, SelectGroup, SelectItem, SelectLabel, SelectSeparator, SelectTrigger, SelectValue,
} from '@/components/ui/select';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { EmptyState, Pill, TableSkeleton } from '@/admin/components';
import { useAdminSenders, useAdminUsers } from '@/admin/queries';
import { formatDateTime, relativeTime } from '@/admin/format';
import type { AuditAction, AuditEntry } from '../../../shared/types';

const ACTION_LABELS: Record<AuditAction, string> = {
  'auth.login': 'Signed in',
  'auth.login_failed': 'Failed sign-in attempt',
  'auth.logout': 'Signed out',
  'auth.password_changed': 'Changed own password',
  'user.create': 'Created user',
  'user.update': 'Edited user',
  'user.disable': 'Disabled user',
  'user.enable': 'Enabled user',
  'user.reset_password': 'Reset user password',
  'user.senders_set': 'Changed sender access',
  'sender.create': 'Added sender account',
  'sender.update': 'Edited sender account',
  'sender.password_rotated': 'Replaced App Password',
  'sender.verify': 'Re-verified sender account',
  'sender.disable': 'Disabled sender account',
  'sender.enable': 'Enabled sender account',
  'template.create': 'Created template',
  'template.update': 'Edited template',
  'template.delete': 'Deleted template',
  'template.share': 'Changed template sharing',
  'job.create': 'Created send job',
  'job.start': 'Started sending',
  'job.pause': 'Paused sending',
  'job.cancel': 'Cancelled send job',
  'job.complete': 'Send job finished',
  'job.retry_failed': 'Retried failed messages',
  'job.test_send': 'Sent a test email',
  'suppression.add': 'Added to suppression list',
  'suppression.remove': 'Removed from suppression list',
};

const CATEGORIES: { prefix: string; label: string }[] = [
  { prefix: 'auth', label: 'Sign-in and passwords' },
  { prefix: 'user', label: 'Users' },
  { prefix: 'sender', label: 'Sender accounts' },
  { prefix: 'template', label: 'Templates' },
  { prefix: 'job', label: 'Sending' },
  { prefix: 'suppression', label: 'Suppression list' },
];

const actionLabel = (a: string) => ACTION_LABELS[a as AuditAction] ?? a;

function actionTone(a: string): 'neutral' | 'success' | 'warning' | 'danger' | 'gold' {
  if (a === 'auth.login_failed' || a.endsWith('.disable') || a.endsWith('.delete') || a === 'job.cancel') return 'danger';
  if (a.includes('password') || a === 'user.senders_set' || a === 'template.share') return 'warning';
  if (a.endsWith('.create') || a.endsWith('.enable') || a === 'job.complete') return 'success';
  if (a.startsWith('sender.')) return 'gold';
  return 'neutral';
}

interface Filters { action: string; actorId: string; from: string; to: string }
const EMPTY: Filters = { action: 'all', actorId: 'all', from: '', to: '' };

/** Builds the query string; dates are local calendar days converted to exact instants. */
function toParams(f: Filters): URLSearchParams {
  const p = new URLSearchParams();
  if (f.action !== 'all') p.set('action', f.action);
  if (f.actorId !== 'all') p.set('actorId', f.actorId);
  if (f.from) p.set('from', new Date(`${f.from}T00:00:00`).toISOString());
  if (f.to) {
    const end = new Date(`${f.to}T00:00:00`);
    end.setDate(end.getDate() + 1);
    p.set('to', end.toISOString());
  }
  return p;
}

export default function AuditPage() {
  const [filters, setFilters] = useState<Filters>(EMPTY);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const usersQ = useAdminUsers();
  const sendersQ = useAdminSenders();
  const userById = useMemo(() => new Map((usersQ.data ?? []).map(u => [u.id, u])), [usersQ.data]);
  const senderById = useMemo(() => new Map((sendersQ.data ?? []).map(s => [s.id, s])), [sendersQ.data]);

  const params = toParams(filters);
  const paramString = params.toString();
  const dateError = filters.from && filters.to && filters.from > filters.to ? 'The start date is after the end date.' : null;

  const q = useInfiniteQuery({
    queryKey: ['admin', 'audit', paramString],
    enabled: !dateError,
    initialPageParam: null as string | null,
    queryFn: async ({ pageParam }) => {
      const p = new URLSearchParams(paramString);
      if (pageParam) p.set('cursor', pageParam);
      return api.get<{ entries: AuditEntry[]; nextCursor: string | null }>(`/admin/audit?${p}`);
    },
    getNextPageParam: (last) => last.nextCursor,
  });
  const entries = q.data?.pages.flatMap(p => p.entries) ?? [];
  const filtered = paramString !== '';

  const set = <K extends keyof Filters>(k: K, v: Filters[K]) => setFilters(f => ({ ...f, [k]: v }));
  const toggle = (id: string) => setExpanded(s => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });

  function targetText(e: AuditEntry): string {
    if (!e.targetType) return '';
    const meta = e.meta ?? {};
    if (e.targetType === 'user' && e.targetId) {
      const u = userById.get(e.targetId);
      return u ? u.name : String(meta.email ?? 'Deleted user');
    }
    if (e.targetType === 'sender_account' && e.targetId) {
      const s = senderById.get(e.targetId);
      return s ? s.displayName : String(meta.email ?? 'Sender account');
    }
    if (e.targetType === 'suppression') return e.targetId ?? '';
    return `${e.targetType.replace(/_/g, ' ')}${e.targetId ? ` ${e.targetId.slice(0, 8)}` : ''}`;
  }

  return (
    <>
      <PageHeader
        title="Audit log"
        description="A permanent record of sign-ins, account changes, template sharing and sending. Entries cannot be edited or deleted."
        actions={
          <Button variant="outline" asChild>
            <a href={`/api/admin/audit/export.csv${paramString ? `?${paramString}` : ''}`} download>
              <Download /> Export CSV
            </a>
          </Button>
        }
      />
      <PageBody className="space-y-4">
        <section aria-label="Filters" className="flex flex-wrap items-end gap-3 rounded-lg border bg-card p-4">
          <div className="w-full space-y-1.5 sm:w-60">
            <Label htmlFor="af-action" className="text-xs">Action</Label>
            <Select value={filters.action} onValueChange={v => set('action', v)}>
              <SelectTrigger id="af-action" className="w-full"><SelectValue /></SelectTrigger>
              <SelectContent className="max-h-80">
                <SelectItem value="all">All actions</SelectItem>
                {CATEGORIES.map(c => (
                  <SelectGroup key={c.prefix}>
                    <SelectSeparator />
                    <SelectLabel>{c.label}</SelectLabel>
                    <SelectItem value={`${c.prefix}.*`}>All {c.label.toLowerCase()} events</SelectItem>
                    {(Object.keys(ACTION_LABELS) as AuditAction[]).filter(a => a.startsWith(c.prefix + '.')).map(a => (
                      <SelectItem key={a} value={a}>{ACTION_LABELS[a]}</SelectItem>
                    ))}
                  </SelectGroup>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="w-full space-y-1.5 sm:w-56">
            <Label htmlFor="af-user" className="text-xs">Performed by</Label>
            <Select value={filters.actorId} onValueChange={v => set('actorId', v)}>
              <SelectTrigger id="af-user" className="w-full"><SelectValue /></SelectTrigger>
              <SelectContent className="max-h-80">
                <SelectItem value="all">Anyone</SelectItem>
                {(usersQ.data ?? []).map(u => <SelectItem key={u.id} value={u.id}>{u.name}</SelectItem>)}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-1.5">
            <Label htmlFor="af-from" className="text-xs">From</Label>
            <Input id="af-from" type="date" value={filters.from} max={filters.to || undefined} onChange={e => set('from', e.target.value)} className="w-40" />
          </div>
          <div className="space-y-1.5">
            <Label htmlFor="af-to" className="text-xs">To</Label>
            <Input id="af-to" type="date" value={filters.to} min={filters.from || undefined} onChange={e => set('to', e.target.value)} className="w-40" />
          </div>
          {filtered && (
            <Button variant="ghost" onClick={() => setFilters(EMPTY)}><X /> Clear filters</Button>
          )}
          {dateError && <p role="alert" className="w-full text-sm text-destructive">{dateError}</p>}
        </section>

        {q.isLoading ? <TableSkeleton rows={8} cols={5} /> : q.isError ? (
          <p role="alert" className="text-sm text-destructive">Could not load the audit log: {errorMessage(q.error)}</p>
        ) : entries.length === 0 ? (
          <EmptyState icon={ScrollText} title={filtered ? 'No entries match these filters' : 'No activity recorded yet'}>
            {filtered ? 'Try a wider date range or a different action.' : 'Sign-ins and changes will appear here as they happen.'}
          </EmptyState>
        ) : (
          <>
            <div className="overflow-hidden rounded-lg border bg-card">
              <Table>
                <TableHeader>
                  <TableRow className="bg-muted/50 hover:bg-muted/50">
                    <TableHead className="w-10 pl-3"><span className="sr-only">Details</span></TableHead>
                    <TableHead>When</TableHead>
                    <TableHead>Who</TableHead>
                    <TableHead>Action</TableHead>
                    <TableHead>Subject</TableHead>
                    <TableHead className="hidden lg:table-cell">IP address</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {entries.map(e => {
                    const open = expanded.has(e.id);
                    const actor = e.actorId ? userById.get(e.actorId) : undefined;
                    const hasMeta = !!e.meta && Object.keys(e.meta).length > 0;
                    return (
                      <Fragment key={e.id}>
                        <TableRow className={cn(open && 'border-b-0 bg-muted/30')}>
                          <TableCell className="pl-3">
                            {hasMeta && (
                              <Button
                                variant="ghost" size="icon-xs" onClick={() => toggle(e.id)}
                                aria-expanded={open} aria-controls={`audit-meta-${e.id}`}
                                aria-label={open ? 'Hide details' : 'Show details'}
                              >
                                {open ? <ChevronDown /> : <ChevronRight />}
                              </Button>
                            )}
                          </TableCell>
                          <TableCell className="whitespace-nowrap text-sm">
                            <div>{formatDateTime(e.at)}</div>
                            <div className="text-xs text-muted-foreground">{relativeTime(e.at)}</div>
                          </TableCell>
                          <TableCell className="text-sm">
                            <div className="font-medium">{actor?.name ?? e.actorEmail ?? 'System'}</div>
                            {actor && <div className="text-xs text-muted-foreground">{actor.email}</div>}
                          </TableCell>
                          <TableCell>
                            <Pill tone={actionTone(e.action)}>{actionLabel(e.action)}</Pill>
                            <div className="mt-1 font-mono text-[11px] text-muted-foreground">{e.action}</div>
                          </TableCell>
                          <TableCell className="max-w-[16rem] truncate text-sm">{targetText(e)}</TableCell>
                          <TableCell className="hidden font-mono text-xs text-muted-foreground lg:table-cell">{e.ip ?? ''}</TableCell>
                        </TableRow>
                        {open && hasMeta && (
                          <TableRow className="bg-muted/30 hover:bg-muted/30">
                            <TableCell />
                            <TableCell colSpan={5} className="pb-4 pt-0">
                              <pre id={`audit-meta-${e.id}`} className="max-h-72 overflow-auto whitespace-pre-wrap break-all rounded-md border bg-card p-3 font-mono text-xs leading-relaxed">
                                {JSON.stringify(e.meta, null, 2)}
                              </pre>
                            </TableCell>
                          </TableRow>
                        )}
                      </Fragment>
                    );
                  })}
                </TableBody>
              </Table>
            </div>
            <div className="flex items-center justify-between gap-3">
              <span className="text-sm text-muted-foreground">Showing {entries.length.toLocaleString('en-IN')} entries, newest first</span>
              {q.hasNextPage ? (
                <Button variant="outline" onClick={() => q.fetchNextPage()} disabled={q.isFetchingNextPage}>
                  {q.isFetchingNextPage ? <><Loader2 className="animate-spin" /> Loading…</> : 'Load more'}
                </Button>
              ) : (
                <span className="text-sm text-muted-foreground">End of the log</span>
              )}
            </div>
          </>
        )}
      </PageBody>
    </>
  );
}
