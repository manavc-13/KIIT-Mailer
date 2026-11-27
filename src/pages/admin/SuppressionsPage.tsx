// Admin: suppression list. Every send job skips these addresses automatically.
import { useMemo, useState, type FormEvent } from 'react';
import { useMutation, useQueryClient } from '@tanstack/react-query';
import { Info, Loader2, Plus, ShieldBan, Trash2 } from 'lucide-react';
import { toast } from 'sonner';
import { PageHeader, PageBody } from '@/components/AppShell';
import { api, errorMessage } from '@/lib/api';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { ConfirmDialog, EmptyState, Pill, SearchBox, TableSkeleton } from '@/admin/components';
import { SUPPRESSIONS_KEY, SUPPRESSIONS_LIST_MAX, useSuppressions, type Suppression } from '@/admin/queries';
import { useDebouncedValue } from '@/preview/useRenderPreview';
import { formatDateTime, relativeTime } from '@/admin/format';

const REASONS: Record<Suppression['reason'], { label: string; tone: 'neutral' | 'warning' | 'danger' }> = {
  unsubscribed: { label: 'Unsubscribed', tone: 'neutral' },
  bounced: { label: 'Bounced', tone: 'danger' },
  manual: { label: 'Added by administrator', tone: 'warning' },
};

export default function SuppressionsPage() {
  const qc = useQueryClient();
  const listQ = useSuppressions();
  const [search, setSearch] = useState('');
  // Server-side search (addresses only) reaches entries beyond the newest 5,000 the plain list returns.
  const term = useDebouncedValue(search.trim().toLowerCase(), 300);
  const searchQ = useSuppressions(term, { enabled: !!term });
  const [email, setEmail] = useState('');
  const [note, setNote] = useState('');
  const [formError, setFormError] = useState<string | null>(null);
  const [removing, setRemoving] = useState<Suppression | null>(null);

  const rows = useMemo(() => {
    const q = search.trim().toLowerCase();
    const all = listQ.data ?? [];
    if (!q) return all;
    // Client filter (address or note) over the loaded list: instant while typing, and the only way to match notes.
    const local = all.filter(s => s.email.includes(q) || (s.note ?? '').toLowerCase().includes(q));
    const server = term && searchQ.data ? searchQ.data.filter(s => s.email.includes(q)) : [];
    if (!server.length) return local;
    const seen = new Set(server.map(s => s.email));
    return [...server, ...local.filter(s => !seen.has(s.email))].sort((a, b) => b.at.localeCompare(a.at));
  }, [listQ.data, search, term, searchQ.data]);
  // The server returns at most 5,000 rows (newest first), with or without a search.
  const capped = ((term ? searchQ.data : listQ.data)?.length ?? 0) >= SUPPRESSIONS_LIST_MAX;

  const add = useMutation({
    mutationFn: (req: { email: string; note: string | null }) =>
      api.post<{ ok: true; alreadySuppressed?: boolean }>('/admin/suppressions', req),
    onSuccess: (res, req) => {
      qc.invalidateQueries({ queryKey: SUPPRESSIONS_KEY });
      if (res.alreadySuppressed) toast.info(`${req.email} was already on the suppression list.`);
      else toast.success(`${req.email} will be skipped by all future sends.`);
      setEmail(''); setNote(''); setFormError(null);
    },
    onError: (err) => setFormError(errorMessage(err)),
  });

  function submit(e: FormEvent) {
    e.preventDefault();
    const v = email.trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v)) return setFormError('Enter a valid email address.');
    setFormError(null);
    add.mutate({ email: v, note: note.trim() || null });
  }

  async function remove(s: Suppression) {
    try {
      await api.del(`/admin/suppressions/${encodeURIComponent(s.email)}`);
      qc.invalidateQueries({ queryKey: SUPPRESSIONS_KEY });
      toast.success(`${s.email} was removed. Future sends will include this address again.`);
    } catch (err) { toast.error(errorMessage(err)); throw err; }
  }

  return (
    <>
      <PageHeader
        title="Suppression list"
        description="Addresses on this list are skipped automatically by every send, from every account. You never need to remove them from recipient lists by hand."
      />
      <PageBody className="space-y-5">
        <div className="grid gap-5 lg:grid-cols-[minmax(0,1fr)_minmax(0,22rem)]">
          <div className="flex gap-3 rounded-lg border bg-card p-4 text-sm text-muted-foreground">
            <Info className="mt-0.5 size-4 shrink-0 text-[var(--gold)]" aria-hidden="true" />
            <div className="space-y-1.5">
              <p>Addresses are added in three ways:</p>
              <ul className="list-disc space-y-1 pl-5">
                <li><strong className="font-medium text-foreground">Unsubscribed</strong>: the recipient used the one-click unsubscribe link in a bulk email.</li>
                <li><strong className="font-medium text-foreground">Bounced</strong>: mail to the address could not be delivered.</li>
                <li><strong className="font-medium text-foreground">Added by administrator</strong>: someone asked not to be emailed, or the address should never receive bulk mail.</li>
              </ul>
              <p>Removing an address means future sends will include it again. Only do this if the person has asked to receive mail again.</p>
            </div>
          </div>

          <form onSubmit={submit} noValidate className="space-y-3 rounded-lg border bg-card p-4" aria-labelledby="add-suppression-title">
            <h2 id="add-suppression-title" className="text-base font-semibold">Add an address</h2>
            {formError && <p role="alert" className="text-sm text-destructive">{formError}</p>}
            <div className="space-y-1.5">
              <Label htmlFor="sup-email">Email address</Label>
              <Input id="sup-email" type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="name@example.com" />
            </div>
            <div className="space-y-1.5">
              <Label htmlFor="sup-note">Note (optional)</Label>
              <Input id="sup-note" value={note} onChange={e => setNote(e.target.value)} maxLength={500} placeholder="e.g. Requested by phone on 12 Nov" />
            </div>
            <Button type="submit" disabled={add.isPending} className="w-full">
              {add.isPending ? <Loader2 className="animate-spin" /> : <Plus />} Add to suppression list
            </Button>
          </form>
        </div>

        {listQ.isLoading ? <TableSkeleton cols={4} /> : listQ.isError ? (
          <p role="alert" className="text-sm text-destructive">Could not load the suppression list: {errorMessage(listQ.error)}</p>
        ) : !listQ.data?.length ? (
          <EmptyState icon={ShieldBan} title="No suppressed addresses">
            When someone unsubscribes or an address bounces, it appears here and is skipped from then on.
          </EmptyState>
        ) : (
          <div className="space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <SearchBox value={search} onChange={setSearch} label="Search suppressed addresses" placeholder="Search address or note" />
              <span className="text-sm text-muted-foreground" aria-live="polite">
                {search.trim()
                  ? `${rows.length.toLocaleString('en-IN')} matching address${rows.length === 1 ? '' : 'es'}`
                  : `${rows.length.toLocaleString('en-IN')} of ${listQ.data.length.toLocaleString('en-IN')} addresses`}
                {!!term && searchQ.isFetching && ' · searching…'}
                {capped && ` · Showing the newest ${SUPPRESSIONS_LIST_MAX.toLocaleString('en-IN')}${term ? ' matches' : '; search to find older ones'}`}
              </span>
            </div>
            <div className="overflow-hidden rounded-lg border bg-card">
              <Table>
                <TableHeader>
                  <TableRow className="bg-muted/50 hover:bg-muted/50">
                    <TableHead className="pl-4">Email address</TableHead>
                    <TableHead>Reason</TableHead>
                    <TableHead>Note</TableHead>
                    <TableHead>Added</TableHead>
                    <TableHead className="w-12 pr-4"><span className="sr-only">Actions</span></TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {rows.length === 0 && (
                    <TableRow><TableCell colSpan={5} className="py-10 text-center text-sm text-muted-foreground">No addresses match “{search}”.</TableCell></TableRow>
                  )}
                  {rows.map(s => (
                    <TableRow key={s.email}>
                      <TableCell className="pl-4 font-medium">{s.email}</TableCell>
                      <TableCell><Pill tone={REASONS[s.reason].tone}>{REASONS[s.reason].label}</Pill></TableCell>
                      <TableCell className="max-w-[20rem] whitespace-normal text-sm text-muted-foreground">{s.note ?? ''}</TableCell>
                      <TableCell className="whitespace-nowrap text-sm" title={formatDateTime(s.at)}>{relativeTime(s.at)}</TableCell>
                      <TableCell className="pr-4 text-right">
                        <Button variant="ghost" size="icon-sm" aria-label={`Remove ${s.email}`} onClick={() => setRemoving(s)}>
                          <Trash2 />
                        </Button>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </div>
        )}
      </PageBody>

      <ConfirmDialog
        open={!!removing} onOpenChange={o => { if (!o) setRemoving(null); }}
        title={`Remove ${removing?.email ?? ''}?`}
        description={<>
          <p>Future sends will include this address again.</p>
          {removing?.reason === 'unsubscribed' && <p className="font-medium text-foreground">This person unsubscribed themselves. Remove them only if they have asked to receive mail again.</p>}
        </>}
        confirmLabel="Remove address"
        destructive
        onConfirm={() => remove(removing!)}
      />
    </>
  );
}
