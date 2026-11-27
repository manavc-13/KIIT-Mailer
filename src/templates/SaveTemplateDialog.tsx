// "Save as template" / "Update template" dialog used by Compose.
import { useEffect, useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { AlertCircle, Loader2 } from 'lucide-react';
import { toast } from 'sonner';
import { api, errorMessage } from '@/lib/api';
import { Button } from '@/components/ui/button';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Label } from '@/components/ui/label';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import {
  Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle,
} from '@/components/ui/dialog';
import type { SaveTemplateRequest, Template } from '../../shared/types';
import { categoriesOf, primeTemplate, useTemplates } from './hooks';
import { NameCategoryFields } from './fields';
import { copyNameFor } from './importExport';

export interface SaveTemplateDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  /** Content to save (without name/category, which the dialog asks for). */
  content: Omit<SaveTemplateRequest, 'name' | 'category'>;
  /** If set and owned by the user, the dialog offers "Update <name>" alongside "Save as new". */
  existing?: { id: string; name: string; category: string | null; isOwner: boolean } | null;
  onSaved: (template: Template) => void;
}

type Mode = 'update' | 'new';

/** "Save as template" / "Update template" dialog used by Compose. */
export function SaveTemplateDialog({ open, onOpenChange, content, existing, onSaved }: SaveTemplateDialogProps) {
  const qc = useQueryClient();
  const { data: list } = useTemplates();
  const canUpdate = !!existing?.isOwner;
  const [mode, setMode] = useState<Mode>('new');
  const [name, setName] = useState('');
  const [category, setCategory] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [nameError, setNameError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  // Reset each time the dialog opens (keyed on primitives so an inline `existing` object does not reset typing).
  const exName = existing?.name ?? '';
  const exCategory = existing?.category ?? '';
  useEffect(() => {
    if (!open) return;
    setMode(canUpdate ? 'update' : 'new');
    setName(exName);
    setCategory(exCategory);
    setError(null); setNameError(null); setSaving(false);
  }, [open, canUpdate, exName, exCategory, existing?.id]);

  function switchMode(m: Mode) {
    setMode(m);
    if (!existing) return;
    // Suggest a distinct name for a new copy of your own template; restore it when switching back.
    if (m === 'new' && name.trim() === existing.name) setName(copyNameFor(existing.name));
    if (m === 'update' && name.trim() === copyNameFor(existing.name)) setName(existing.name);
  }

  async function save(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    const trimmed = name.trim();
    if (!trimmed) { setNameError('Give the template a name.'); return; }
    setNameError(null);
    const req: SaveTemplateRequest = { ...content, name: trimmed, category: category.trim() || null };
    setSaving(true);
    try {
      const { template } = mode === 'update' && existing
        ? await api.put<{ template: Template }>(`/templates/${existing.id}`, req)
        : await api.post<{ template: Template }>('/templates', req);
      primeTemplate(qc, template);
      await qc.invalidateQueries({ queryKey: ['templates'] });
      toast.success(mode === 'update' ? `Updated "${template.name}"` : `Saved "${template.name}" to My templates`);
      onSaved(template);
      onOpenChange(false);
    } catch (err) {
      setError(errorMessage(err));
    } finally {
      setSaving(false);
    }
  }

  return (
    <Dialog open={open} onOpenChange={(o) => !saving && onOpenChange(o)}>
      <DialogContent className="sm:max-w-md">
        <form onSubmit={save} className="grid gap-5" noValidate>
          <DialogHeader>
            <DialogTitle className="font-serif text-xl">{canUpdate ? 'Save template' : 'Save as template'}</DialogTitle>
            <DialogDescription>
              {canUpdate
                ? 'Update the template you started from, or keep it unchanged and save a new one.'
                : 'Templates are private to you until an administrator shares them.'}
            </DialogDescription>
          </DialogHeader>

          {canUpdate && existing && (
            <RadioGroup value={mode} onValueChange={v => switchMode(v as Mode)} className="gap-2">
              <Label className="flex cursor-pointer items-start gap-3 rounded-md border p-3 font-normal has-[[data-state=checked]]:border-[var(--gold)] has-[[data-state=checked]]:bg-[var(--gold-soft)]">
                <RadioGroupItem value="update" className="mt-0.5" />
                <span className="min-w-0">
                  <span className="block truncate font-medium">Update “{existing.name}”</span>
                  <span className="block text-xs text-muted-foreground">Replaces the saved content of this template.</span>
                </span>
              </Label>
              <Label className="flex cursor-pointer items-start gap-3 rounded-md border p-3 font-normal has-[[data-state=checked]]:border-[var(--gold)] has-[[data-state=checked]]:bg-[var(--gold-soft)]">
                <RadioGroupItem value="new" className="mt-0.5" />
                <span>
                  <span className="block font-medium">Save as a new template</span>
                  <span className="block text-xs text-muted-foreground">The original stays as it is.</span>
                </span>
              </Label>
            </RadioGroup>
          )}

          <NameCategoryFields
            name={name} onName={v => { setName(v); if (nameError) setNameError(null); }}
            category={category} onCategory={setCategory}
            suggestions={categoriesOf(list)} nameError={nameError}
            autoFocus={!canUpdate}
          />

          {error && (
            <Alert variant="destructive">
              <AlertCircle />
              <AlertDescription>{error}</AlertDescription>
            </Alert>
          )}

          <DialogFooter>
            <Button type="button" variant="outline" onClick={() => onOpenChange(false)} disabled={saving}>Cancel</Button>
            <Button type="submit" disabled={saving}>
              {saving && <Loader2 className="animate-spin" />}
              {mode === 'update' ? 'Update template' : 'Save template'}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
