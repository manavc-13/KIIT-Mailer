// Dialogs used by the template library: rename/category, share (admin), delete, import.
import { useEffect, useMemo, useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { AlertCircle, CheckCircle2, Loader2, Search } from 'lucide-react';
import { toast } from 'sonner';
import { api, errorMessage } from '@/lib/api';
import { Button } from '@/components/ui/button';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Checkbox } from '@/components/ui/checkbox';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import {
  Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle,
} from '@/components/ui/dialog';
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription,
  AlertDialogFooter, AlertDialogHeader, AlertDialogTitle,
} from '@/components/ui/alert-dialog';
import type { SaveTemplateRequest, Template, TemplateSummary, TemplateVisibility } from '../../shared/types';
import { categoriesOf, fetchTemplate, primeTemplate, themeLabel, useAdminUsers, useTemplates } from './hooks';
import { NameCategoryFields } from './fields';
import { chunk, LEGACY_STORAGE_KEY, templateToRequest, type ImportFormat } from './importExport';

function ErrorAlert({ message }: { message: string | null }) {
  if (!message) return null;
  return (
    <Alert variant="destructive">
      <AlertCircle />
      <AlertDescription>{message}</AlertDescription>
    </Alert>
  );
}

const optionClass =
  'flex cursor-pointer items-start gap-3 rounded-md border p-3 font-normal has-[[data-state=checked]]:border-[var(--gold)] has-[[data-state=checked]]:bg-[var(--gold-soft)]';

// ------------------------------------------------------------------ rename / category

export function RenameDialog({ target, onOpenChange }: {
  target: TemplateSummary | null; onOpenChange: (open: boolean) => void;
}) {
  const qc = useQueryClient();
  const { data: list } = useTemplates();
  const [name, setName] = useState('');
  const [category, setCategory] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (!target) return;
    setName(target.name); setCategory(target.category ?? ''); setError(null); setSaving(false);
  }, [target]);

  async function save(e: React.FormEvent) {
    e.preventDefault();
    if (!target) return;
    if (!name.trim()) { setError('Give the template a name.'); return; }
    setSaving(true); setError(null);
    try {
      // PUT needs the whole template, so fetch the current content and send it back with the new name/category.
      const full = await fetchTemplate(qc, target);
      const req: SaveTemplateRequest = { ...templateToRequest(full), name: name.trim(), category: category.trim() || null };
      const { template } = await api.put<{ template: Template }>(`/templates/${target.id}`, req);
      primeTemplate(qc, template);
      await qc.invalidateQueries({ queryKey: ['templates'] });
      toast.success('Template details saved');
      onOpenChange(false);
    } catch (err) {
      setError(errorMessage(err));
    } finally {
      setSaving(false);
    }
  }

  return (
    <Dialog open={!!target} onOpenChange={o => !saving && onOpenChange(o)}>
      <DialogContent className="sm:max-w-md">
        <form onSubmit={save} className="grid gap-5" noValidate>
          <DialogHeader>
            <DialogTitle className="font-serif text-xl">Rename or recategorise</DialogTitle>
            <DialogDescription>The content of the template is not changed.</DialogDescription>
          </DialogHeader>
          <NameCategoryFields
            name={name} onName={setName} category={category} onCategory={setCategory}
            suggestions={categoriesOf(list)}
          />
          <ErrorAlert message={error} />
          <DialogFooter>
            <Button type="button" variant="outline" onClick={() => onOpenChange(false)} disabled={saving}>Cancel</Button>
            <Button type="submit" disabled={saving}>{saving && <Loader2 className="animate-spin" />}Save</Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

// ------------------------------------------------------------------ share (admin)

export function ShareDialog({ target, onOpenChange }: {
  target: TemplateSummary | null; onOpenChange: (open: boolean) => void;
}) {
  const qc = useQueryClient();
  const open = !!target;
  const users = useAdminUsers(open);
  const [visibility, setVisibility] = useState<TemplateVisibility>('private');
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [filter, setFilter] = useState('');
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!target) return;
    let cancelled = false;
    setVisibility(target.visibility); setSelected(new Set()); setFilter(''); setError(null); setSaving(false);
    setLoading(true);
    fetchTemplate(qc, target)
      .then(full => { if (!cancelled) setSelected(new Set(full.sharedUserIds)); })
      .catch(err => { if (!cancelled) setError(errorMessage(err)); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [target, qc]);

  const candidates = useMemo(() => {
    const q = filter.trim().toLowerCase();
    return (users.data ?? [])
      .filter(u => u.id !== target?.ownerId)
      .filter(u => !q || u.name.toLowerCase().includes(q) || u.email.toLowerCase().includes(q))
      .sort((a, b) => Number(selected.has(b.id)) - Number(selected.has(a.id)) || a.name.localeCompare(b.name));
    // `selected` deliberately excluded: re-sorting while ticking boxes would move rows under the pointer.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [users.data, filter, target?.ownerId]);

  function toggle(id: string, on: boolean) {
    setSelected(prev => { const n = new Set(prev); if (on) n.add(id); else n.delete(id); return n; });
  }

  async function save() {
    if (!target) return;
    if (visibility === 'shared_users' && selected.size === 0) { setError('Choose at least one person to share with.'); return; }
    setSaving(true); setError(null);
    try {
      const { template } = await api.put<{ template: Template }>(`/templates/${target.id}/share`, {
        visibility, userIds: visibility === 'shared_users' ? [...selected] : [],
      });
      primeTemplate(qc, template);
      await qc.invalidateQueries({ queryKey: ['templates'] });
      toast.success(
        visibility === 'private' ? `"${target.name}" is now private`
          : visibility === 'shared_all' ? `"${target.name}" is shared with everyone`
            : `"${target.name}" is shared with ${selected.size} ${selected.size === 1 ? 'person' : 'people'}`,
      );
      onOpenChange(false);
    } catch (err) {
      setError(errorMessage(err));
    } finally {
      setSaving(false);
    }
  }

  return (
    <Dialog open={open} onOpenChange={o => !saving && onOpenChange(o)}>
      <DialogContent className="sm:max-w-lg">
        <DialogHeader>
          <DialogTitle className="font-serif text-xl">Share “{target?.name}”</DialogTitle>
          <DialogDescription>
            People it is shared with can use and duplicate it, but only the owner ({target?.isOwner ? 'you' : target?.ownerName}) can edit it.
          </DialogDescription>
        </DialogHeader>

        <RadioGroup value={visibility} onValueChange={v => setVisibility(v as TemplateVisibility)} className="gap-2">
          <Label className={optionClass}>
            <RadioGroupItem value="private" className="mt-0.5" />
            <span><span className="block font-medium">Private</span><span className="block text-xs text-muted-foreground">Only the owner can see it.</span></span>
          </Label>
          <Label className={optionClass}>
            <RadioGroupItem value="shared_all" className="mt-0.5" />
            <span><span className="block font-medium">Everyone</span><span className="block text-xs text-muted-foreground">Every KIIT Mailer user sees it under “Shared with me”.</span></span>
          </Label>
          <Label className={optionClass}>
            <RadioGroupItem value="shared_users" className="mt-0.5" />
            <span><span className="block font-medium">Specific people</span><span className="block text-xs text-muted-foreground">Only the people you choose below.</span></span>
          </Label>
        </RadioGroup>

        {visibility === 'shared_users' && (
          <div className="grid gap-2">
            <div className="flex items-center justify-between">
              <Label htmlFor="share-filter">People</Label>
              <span className="text-xs text-muted-foreground">{selected.size} selected</span>
            </div>
            <div className="relative">
              <Search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
              <Input id="share-filter" value={filter} onChange={e => setFilter(e.target.value)} placeholder="Search by name or email" className="pl-8" />
            </div>
            <div className="max-h-60 overflow-y-auto rounded-md border" role="group" aria-label="People to share with">
              {users.isLoading || loading ? (
                <div className="flex items-center justify-center gap-2 py-6 text-sm text-muted-foreground"><Loader2 className="size-4 animate-spin" />Loading…</div>
              ) : users.isError ? (
                <p className="p-4 text-sm text-destructive">Could not load users: {errorMessage(users.error)}</p>
              ) : candidates.length === 0 ? (
                <p className="p-4 text-sm text-muted-foreground">{filter ? 'No one matches.' : 'There are no other users yet.'}</p>
              ) : candidates.map(u => (
                <label key={u.id} className="flex cursor-pointer items-center gap-3 border-b px-3 py-2 last:border-b-0 hover:bg-muted/60">
                  <Checkbox checked={selected.has(u.id)} onCheckedChange={v => toggle(u.id, v === true)} aria-label={`Share with ${u.name}`} />
                  <span className="min-w-0 flex-1">
                    <span className="block truncate text-sm font-medium">{u.name}{u.disabled && <span className="ml-1.5 text-xs font-normal text-muted-foreground">(disabled)</span>}</span>
                    <span className="block truncate text-xs text-muted-foreground">{u.email}</span>
                  </span>
                  {u.role === 'admin' && <span className="text-[11px] uppercase tracking-wider text-muted-foreground">Admin</span>}
                </label>
              ))}
            </div>
          </div>
        )}

        <ErrorAlert message={error} />
        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)} disabled={saving}>Cancel</Button>
          <Button onClick={save} disabled={saving || loading}>{saving && <Loader2 className="animate-spin" />}Save sharing</Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

// ------------------------------------------------------------------ delete

export function DeleteDialog({ target, onOpenChange }: {
  target: TemplateSummary | null; onOpenChange: (open: boolean) => void;
}) {
  const qc = useQueryClient();
  const [busy, setBusy] = useState(false);

  async function confirm(e: React.MouseEvent) {
    e.preventDefault();
    if (!target) return;
    setBusy(true);
    try {
      await api.del(`/templates/${target.id}`);
      await qc.invalidateQueries({ queryKey: ['templates'] });
      toast.success(`Deleted "${target.name}"`);
      onOpenChange(false);
    } catch (err) {
      toast.error(errorMessage(err));
    } finally {
      setBusy(false);
    }
  }

  const sharedNote = target && target.visibility !== 'private'
    ? ' It is shared, so it will also disappear for everyone it is shared with.' : '';
  return (
    <AlertDialog open={!!target} onOpenChange={o => !busy && onOpenChange(o)}>
      <AlertDialogContent>
        <AlertDialogHeader>
          <AlertDialogTitle className="font-serif">Delete “{target?.name}”?</AlertDialogTitle>
          <AlertDialogDescription>
            {target && !target.isOwner ? `This template belongs to ${target.ownerName}. ` : ''}
            This cannot be undone.{sharedNote} Past mail sent with it is not affected.
          </AlertDialogDescription>
        </AlertDialogHeader>
        <AlertDialogFooter>
          <AlertDialogCancel disabled={busy}>Cancel</AlertDialogCancel>
          <AlertDialogAction variant="destructive" onClick={confirm} disabled={busy}>
            {busy && <Loader2 className="animate-spin" />}Delete template
          </AlertDialogAction>
        </AlertDialogFooter>
      </AlertDialogContent>
    </AlertDialog>
  );
}

// ------------------------------------------------------------------ import

export interface PendingImport {
  source: 'file' | 'browser';
  format: ImportFormat;
  fileName?: string;
  templates: SaveTemplateRequest[];
}

export function ImportDialog({ pending, onOpenChange }: {
  pending: PendingImport | null; onOpenChange: (open: boolean) => void;
}) {
  const qc = useQueryClient();
  const [checked, setChecked] = useState<Set<number>>(new Set());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [done, setDone] = useState<{ created: number; renamed: number } | null>(null);
  const [removed, setRemoved] = useState(false);

  useEffect(() => {
    if (!pending) return;
    setChecked(new Set(pending.templates.map((_, i) => i)));
    setBusy(false); setError(null); setDone(null); setRemoved(false);
  }, [pending]);

  const items = pending?.templates ?? [];
  const allOn = checked.size === items.length && items.length > 0;

  async function run() {
    const chosen = items.filter((_, i) => checked.has(i));
    if (!chosen.length) { setError('Select at least one template.'); return; }
    setBusy(true); setError(null);
    let created = 0, renamed = 0;
    try {
      for (const part of chunk(chosen, 100)) {
        const r = await api.post<{ created: number; renamed?: number }>('/templates/import', { templates: part });
        created += r.created; renamed += r.renamed ?? 0;
      }
      setDone({ created, renamed });
      toast.success(`Imported ${created} template${created === 1 ? '' : 's'}`);
    } catch (err) {
      setError(created ? `Imported ${created}, then stopped: ${errorMessage(err)}` : errorMessage(err));
      if (created) setDone({ created, renamed });
    } finally {
      setBusy(false);
      await qc.invalidateQueries({ queryKey: ['templates'] });
    }
  }

  function removeLegacy() {
    try { localStorage.removeItem(LEGACY_STORAGE_KEY); setRemoved(true); toast.success('Removed the old v1 templates from this browser'); }
    catch (err) { toast.error(`Could not remove them: ${errorMessage(err)}`); }
  }

  const title = pending?.source === 'browser' ? 'Import from this browser (KIIT Mailer v1)' : 'Import templates';
  return (
    <Dialog open={!!pending} onOpenChange={o => !busy && onOpenChange(o)}>
      <DialogContent className="sm:max-w-lg">
        <DialogHeader>
          <DialogTitle className="font-serif text-xl">{title}</DialogTitle>
          <DialogDescription>
            {done ? 'Import finished.' : (
              <>
                {pending?.fileName ? <>From <span className="font-medium text-foreground">{pending.fileName}</span>. </> : null}
                {pending?.format === 'legacy'
                  ? 'These are KIIT Mailer v1 templates. They are imported with the Plain theme so their original styling is kept, into the category “Imported (v1)”.'
                  : 'Templates are added to My templates as private copies.'}
                {' '}Names you already use get “(imported)” added.
              </>
            )}
          </DialogDescription>
        </DialogHeader>

        {done ? (
          <div className="grid gap-4">
            <div className="flex items-start gap-3 rounded-md border bg-[var(--success-soft)] p-3 text-sm">
              <CheckCircle2 className="mt-0.5 size-4 shrink-0 text-[var(--success)]" />
              <span>
                {done.created} template{done.created === 1 ? '' : 's'} added to My templates
                {done.renamed ? `; ${done.renamed} renamed to avoid duplicate names` : ''}.
              </span>
            </div>
            {pending?.source === 'browser' && (
              <div className="rounded-md border p-3 text-sm">
                {removed ? (
                  <p className="text-muted-foreground">The old copies have been removed from this browser.</p>
                ) : (
                  <>
                    <p>The old v1 copies are still stored in this browser. Remove them now that they are in your account?</p>
                    <div className="mt-3 flex gap-2">
                      <Button size="sm" variant="outline" onClick={removeLegacy}>Remove old copies</Button>
                      <Button size="sm" variant="ghost" onClick={() => onOpenChange(false)}>Keep them</Button>
                    </div>
                  </>
                )}
              </div>
            )}
            <ErrorAlert message={error} />
            <DialogFooter><Button onClick={() => onOpenChange(false)}>Done</Button></DialogFooter>
          </div>
        ) : (
          <>
            <div className="rounded-md border">
              <label className="flex cursor-pointer items-center gap-3 border-b bg-muted/50 px-3 py-2 text-sm font-medium">
                <Checkbox
                  checked={allOn ? true : checked.size ? 'indeterminate' : false}
                  onCheckedChange={v => setChecked(v === true ? new Set(items.map((_, i) => i)) : new Set())}
                  aria-label="Select all"
                />
                {items.length} template{items.length === 1 ? '' : 's'} found
                <span className="ml-auto text-xs font-normal text-muted-foreground">{checked.size} selected</span>
              </label>
              <div className="max-h-64 overflow-y-auto">
                {items.map((t, i) => (
                  <label key={i} className="flex cursor-pointer items-center gap-3 border-b px-3 py-2 last:border-b-0 hover:bg-muted/60">
                    <Checkbox
                      checked={checked.has(i)}
                      onCheckedChange={v => setChecked(prev => { const n = new Set(prev); if (v === true) n.add(i); else n.delete(i); return n; })}
                      aria-label={`Import ${t.name}`}
                    />
                    <span className="min-w-0 flex-1">
                      <span className="block truncate text-sm font-medium">{t.name}</span>
                      <span className="block truncate text-xs text-muted-foreground">
                        {[t.subject || 'No subject', t.editorMode === 'html' ? 'HTML' : 'Rich text', themeLabel(t.theme)].join(' · ')}
                      </span>
                    </span>
                  </label>
                ))}
              </div>
            </div>
            <ErrorAlert message={error} />
            <DialogFooter>
              <Button variant="outline" onClick={() => onOpenChange(false)} disabled={busy}>Cancel</Button>
              <Button onClick={run} disabled={busy || checked.size === 0}>
                {busy && <Loader2 className="animate-spin" />}Import {checked.size || ''}
              </Button>
            </DialogFooter>
          </>
        )}
      </DialogContent>
    </Dialog>
  );
}
