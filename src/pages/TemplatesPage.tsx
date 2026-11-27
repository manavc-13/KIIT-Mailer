// Template library: personal and shared templates, import/export, sharing (admin).
import { useMemo, useRef, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useQueryClient } from '@tanstack/react-query';
import {
  ChevronDown, Download, FileJson, FilePlus2, LibraryBig, Loader2, MonitorDown, Plus, Search, Upload, X,
} from 'lucide-react';
import { toast } from 'sonner';
import { PageBody, PageHeader } from '@/components/AppShell';
import { api, errorMessage } from '@/lib/api';
import { useAuth } from '@/lib/auth';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Skeleton } from '@/components/ui/skeleton';
import { Tabs, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import type { Template, TemplateSummary } from '../../shared/types';
import {
  TemplateCard, RenameDialog, ShareDialog, DeleteDialog, ImportDialog, type PendingImport,
  buildExport, parseImport, parseImportText, LEGACY_STORAGE_KEY,
  categoriesOf, fetchTemplate, primeTemplate, useAllTemplates, useTemplates,
} from '@/templates';

type TabId = 'all' | 'mine' | 'shared' | 'sharedByMe' | 'everyone';
const ALL = '__all__';
const NONE = '__none__';

async function mapLimit<T, R>(items: T[], limit: number, fn: (t: T) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) { const i = next++; out[i] = await fn(items[i]); }
  }));
  return out;
}

function download(filename: string, text: string) {
  const url = URL.createObjectURL(new Blob([text], { type: 'application/json' }));
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function EmptyState({ tab, filtered, onNew, onImportBrowser, onClear }: {
  tab: TabId; filtered: boolean; onNew: () => void; onImportBrowser: () => void; onClear: () => void;
}) {
  let title: string; let body: React.ReactNode; let actions: React.ReactNode = null;
  if (filtered) {
    title = 'No templates match';
    body = 'Try a different search or category.';
    actions = <Button variant="outline" onClick={onClear}><X /> Clear filters</Button>;
  } else if (tab === 'shared') {
    title = 'Nothing shared with you yet';
    body = 'When an administrator shares a template with you or with everyone, it appears here. You can use it as it is or duplicate it to make your own version.';
  } else if (tab === 'sharedByMe') {
    title = 'You have not shared any templates';
    body = 'Open the menu on one of your templates and choose “Share…” to make it available to everyone or to chosen people.';
  } else {
    title = 'No templates yet';
    body = (
      <>
        Write a mail in Compose and choose <span className="font-medium text-foreground">Save as template</span> to reuse it later.
        If you used the previous KIIT Mailer in this browser, you can bring those templates across.
      </>
    );
    actions = (
      <>
        <Button onClick={onNew}><Plus /> New template</Button>
        <Button variant="outline" onClick={onImportBrowser}><MonitorDown /> Import from this browser</Button>
      </>
    );
  }
  return (
    <div className="mx-auto flex max-w-lg flex-col items-center rounded-lg border border-dashed bg-card px-6 py-14 text-center">
      <span className="mb-4 grid size-12 place-items-center rounded-full bg-[var(--gold-soft)] text-[var(--gold)]">
        <LibraryBig className="size-5" />
      </span>
      <h2 className="text-xl font-semibold">{title}</h2>
      <p className="mt-2 text-sm text-muted-foreground">{body}</p>
      {actions && <div className="mt-6 flex flex-wrap justify-center gap-2">{actions}</div>}
    </div>
  );
}

export default function TemplatesPage() {
  const navigate = useNavigate();
  const qc = useQueryClient();
  const { isAdmin } = useAuth();
  const [tab, setTab] = useState<TabId>('all');
  const [search, setSearch] = useState('');
  const [category, setCategory] = useState<string>(ALL);
  const mineQ = useTemplates();
  const everyoneQ = useAllTemplates(isAdmin && tab === 'everyone');
  const q = tab === 'everyone' ? everyoneQ : mineQ;
  const list = q.data;

  const [renaming, setRenaming] = useState<TemplateSummary | null>(null);
  const [sharing, setSharing] = useState<TemplateSummary | null>(null);
  const [deleting, setDeleting] = useState<TemplateSummary | null>(null);
  const [pendingImport, setPendingImport] = useState<PendingImport | null>(null);
  const [busyId, setBusyId] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const counts = useMemo(() => {
    const l = mineQ.data ?? [];
    return {
      all: l.length,
      mine: l.filter(t => t.isOwner).length,
      shared: l.filter(t => !t.isOwner).length,
      sharedByMe: l.filter(t => t.isOwner && t.visibility !== 'private').length,
    };
  }, [mineQ.data]);

  const categories = useMemo(() => categoriesOf(list), [list]);
  const hasUncategorised = (list ?? []).some(t => !t.category?.trim());

  const shown = useMemo(() => {
    const needle = search.trim().toLowerCase();
    return (list ?? []).filter((t) => {
      if (tab === 'mine' && !t.isOwner) return false;
      if (tab === 'shared' && t.isOwner) return false;
      if (tab === 'sharedByMe' && !(t.isOwner && t.visibility !== 'private')) return false;
      if (category === NONE && t.category?.trim()) return false;
      if (category !== ALL && category !== NONE && (t.category ?? '').trim().toLowerCase() !== category.toLowerCase()) return false;
      if (!needle) return true;
      return [t.name, t.subject, t.category ?? '', t.ownerName].some(s => s.toLowerCase().includes(needle));
    });
  }, [list, tab, search, category]);

  const filtered = !!search.trim() || category !== ALL;

  // ------------------------------------------------------------ actions
  const onUse = (t: TemplateSummary) => navigate(`/compose?template=${t.id}`);
  const onNew = () => navigate('/compose?new=1', { state: { newTemplate: true } });

  async function onDuplicate(t: TemplateSummary) {
    setBusyId(t.id);
    try {
      const { template } = await api.post<{ template: Template }>(`/templates/${t.id}/duplicate`);
      primeTemplate(qc, template);
      await qc.invalidateQueries({ queryKey: ['templates'] });
      toast.success(`Copied to My templates as "${template.name}"`, {
        action: { label: 'Open', onClick: () => navigate(`/compose?template=${template.id}`) },
      });
    } catch (err) {
      toast.error(errorMessage(err));
    } finally {
      setBusyId(null);
    }
  }

  async function onExport() {
    const mine = (mineQ.data ?? []).filter(t => t.isOwner);
    if (!mine.length) { toast.info('You have no templates of your own to export yet.'); return; }
    setExporting(true);
    try {
      const full = await mapLimit(mine, 6, t => fetchTemplate(qc, t));
      const stamp = new Date().toISOString().slice(0, 10);
      download(`kiit-mailer-templates-${stamp}.json`, JSON.stringify(buildExport(full), null, 2));
      toast.success(`Exported ${full.length} template${full.length === 1 ? '' : 's'}`);
    } catch (err) {
      toast.error(`Export failed: ${errorMessage(err)}`);
    } finally {
      setExporting(false);
    }
  }

  async function onFileChosen(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    if (file.size > 30 * 1024 * 1024) { toast.error('That file is too large to be a template export.'); return; }
    try {
      const { format, templates } = parseImportText(await file.text());
      if (!templates.length) { toast.info('That file contains no templates.'); return; }
      setPendingImport({ source: 'file', format, fileName: file.name, templates });
    } catch (err) {
      toast.error(errorMessage(err));
    }
  }

  function onImportBrowser() {
    let raw: string | null = null;
    try { raw = localStorage.getItem(LEGACY_STORAGE_KEY); } catch { /* storage blocked */ }
    if (!raw) {
      toast.info('No KIIT Mailer v1 templates were found in this browser.', {
        description: 'They are only visible from the browser and web address where you used the old version. You can also import a file made with its “Export all” button.',
      });
      return;
    }
    try {
      const { format, templates } = parseImport(JSON.parse(raw));
      if (!templates.length) { toast.info('The old KIIT Mailer storage in this browser is empty.'); return; }
      setPendingImport({ source: 'browser', format, templates });
    } catch (err) {
      toast.error(`Could not read the old templates: ${errorMessage(err)}`);
    }
  }

  const clearFilters = () => { setSearch(''); setCategory(ALL); };

  return (
    <>
      <PageHeader
        title="Templates"
        description="Reusable mail for you and your colleagues. Shared templates can be used or duplicated; only their owner can edit them."
        actions={(
          <>
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <Button variant="outline"><Upload /> Import <ChevronDown className="opacity-60" /></Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="w-64">
                <DropdownMenuItem onSelect={() => fileRef.current?.click()}><FileJson /> From a JSON file</DropdownMenuItem>
                <DropdownMenuItem onSelect={onImportBrowser}><MonitorDown /> From this browser (KIIT Mailer v1)</DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
            <Button variant="outline" onClick={onExport} disabled={exporting || !counts.mine} title={counts.mine ? 'Download your templates as JSON' : 'You have no templates to export'}>
              {exporting ? <Loader2 className="animate-spin" /> : <Download />} Export
            </Button>
            <Button onClick={onNew}><FilePlus2 /> New template</Button>
            <input ref={fileRef} type="file" accept="application/json,.json" className="hidden" onChange={onFileChosen} aria-hidden="true" tabIndex={-1} />
          </>
        )}
      />
      <PageBody>
        <div className="mb-6 flex flex-col gap-3 lg:flex-row lg:items-center">
          <Tabs value={tab} onValueChange={v => setTab(v as TabId)} className="min-w-0 shrink-0 overflow-x-auto overflow-y-hidden">
            <TabsList>
              <TabsTrigger value="all">All <span className="text-xs text-muted-foreground">{counts.all}</span></TabsTrigger>
              <TabsTrigger value="mine">My templates <span className="text-xs text-muted-foreground">{counts.mine}</span></TabsTrigger>
              <TabsTrigger value="shared">Shared with me <span className="text-xs text-muted-foreground">{counts.shared}</span></TabsTrigger>
              {isAdmin && <TabsTrigger value="sharedByMe">Shared by me <span className="text-xs text-muted-foreground">{counts.sharedByMe}</span></TabsTrigger>}
              {isAdmin && <TabsTrigger value="everyone" title="Every template in the system, including other users' private ones">All users</TabsTrigger>}
            </TabsList>
          </Tabs>
          <div className="flex flex-1 flex-col gap-2 sm:flex-row lg:justify-end">
            <div className="relative sm:w-72">
              <Search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
              <Input
                type="search"
                value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="Search name, subject or owner"
                aria-label="Search templates"
                className="pl-8"
              />
            </div>
            <Select value={category} onValueChange={setCategory}>
              <SelectTrigger className="sm:w-52" aria-label="Filter by category">
                <SelectValue placeholder="All categories" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value={ALL}>All categories</SelectItem>
                {categories.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
                {hasUncategorised && <SelectItem value={NONE}>Uncategorised</SelectItem>}
              </SelectContent>
            </Select>
          </div>
        </div>

        {tab === 'everyone' && (
          <p className="mb-4 rounded-md border border-[color-mix(in_oklab,var(--gold)_40%,var(--border))] bg-[var(--gold-soft)] px-3 py-2 text-sm text-[var(--accent-foreground)]">
            Administrator view: every template in the system, including other users’ private ones. You can share or delete them, but only their owners can edit them.
          </p>
        )}

        {q.isLoading ? (
          <div className="grid grid-cols-[repeat(auto-fill,minmax(260px,1fr))] gap-5" aria-busy="true" aria-label="Loading templates">
            {Array.from({ length: 6 }).map((_, i) => (
              <div key={i} className="overflow-hidden rounded-lg border bg-card">
                <Skeleton className="aspect-[4/3] rounded-none" />
                <div className="space-y-2 p-4"><Skeleton className="h-4 w-2/3" /><Skeleton className="h-3 w-full" /><Skeleton className="h-3 w-1/2" /></div>
              </div>
            ))}
          </div>
        ) : q.isError ? (
          <div className="rounded-lg border border-destructive/40 bg-card p-6 text-sm">
            <p className="font-medium text-destructive">Could not load templates.</p>
            <p className="mt-1 text-muted-foreground">{errorMessage(q.error)}</p>
            <Button className="mt-4" variant="outline" onClick={() => q.refetch()}>Try again</Button>
          </div>
        ) : shown.length === 0 ? (
          <EmptyState tab={tab} filtered={filtered} onNew={onNew} onImportBrowser={onImportBrowser} onClear={clearFilters} />
        ) : (
          <>
            <p className="sr-only" aria-live="polite">{shown.length} templates shown</p>
            <div className="grid grid-cols-[repeat(auto-fill,minmax(260px,1fr))] gap-5">
              {shown.map(t => (
                <TemplateCard
                  key={t.id} t={t} isAdmin={isAdmin} busy={busyId === t.id}
                  onUse={onUse} onDuplicate={onDuplicate} onRename={setRenaming} onShare={setSharing} onDelete={setDeleting}
                />
              ))}
            </div>
          </>
        )}
      </PageBody>

      <RenameDialog target={renaming} onOpenChange={o => !o && setRenaming(null)} />
      <ShareDialog target={sharing} onOpenChange={o => !o && setSharing(null)} />
      <DeleteDialog target={deleting} onOpenChange={o => !o && setDeleting(null)} />
      <ImportDialog pending={pendingImport} onOpenChange={o => !o && setPendingImport(null)} />
    </>
  );
}
