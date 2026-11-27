// Searchable template switcher for Compose: Popover + Command, grouped by ownership.
import { useRef, useState } from 'react';
import { Link } from 'react-router-dom';
import { useQueryClient } from '@tanstack/react-query';
import { Check, ChevronsUpDown, LibraryBig, Loader2, Settings2 } from 'lucide-react';
import { toast } from 'sonner';
import { errorMessage } from '@/lib/api';
import { cn } from '@/lib/utils';
import { Button } from '@/components/ui/button';
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover';
import {
  Command, CommandEmpty, CommandGroup, CommandInput, CommandItem, CommandList, CommandSeparator,
} from '@/components/ui/command';
import type { Template, TemplateSummary } from '../../shared/types';
import { fetchTemplate, themeLabel, useTemplates } from './hooks';

export interface TemplatePickerProps {
  /** Currently loaded template id (null = unsaved / scratch). */
  value: string | null;
  /** Called with the full template after the user picks one (the caller confirms overwriting unsaved work). */
  onPick: (template: Template) => void;
  className?: string;
  /** Optional: disable the trigger (e.g. while sending). */
  disabled?: boolean;
}

function Item({ t, selected, loading, onSelect }: {
  t: TemplateSummary; selected: boolean; loading: boolean; onSelect: () => void;
}) {
  return (
    <CommandItem
      value={`${t.name} ${t.category ?? ''} ${t.ownerName} ${t.id}`}
      onSelect={onSelect}
      className="items-start gap-2 py-2"
    >
      <span className="mt-0.5 grid size-4 place-items-center">
        {loading ? <Loader2 className="animate-spin" /> : selected ? <Check className="text-foreground" /> : null}
      </span>
      <span className="min-w-0 flex-1">
        <span className="block truncate font-medium">{t.name}</span>
        <span className="block truncate text-xs text-muted-foreground">
          {[t.category, themeLabel(t.theme), t.isOwner ? null : `by ${t.ownerName}`].filter(Boolean).join(' · ')}
        </span>
      </span>
    </CommandItem>
  );
}

/** Searchable dropdown of the user's own + shared templates, grouped (My templates / Shared with me). */
export function TemplatePicker({ value, onPick, className, disabled }: TemplatePickerProps) {
  const qc = useQueryClient();
  const { data, isLoading, isError } = useTemplates();
  const [open, setOpen] = useState(false);
  const [loadingId, setLoadingId] = useState<string | null>(null);
  /** The most recent pick; a slower fetch for an earlier pick must not overwrite it. */
  const latestPick = useRef(0);

  const current = data?.find(t => t.id === value);
  const mine = (data ?? []).filter(t => t.isOwner);
  const shared = (data ?? []).filter(t => !t.isOwner);

  async function pick(t: TemplateSummary) {
    const seq = ++latestPick.current;
    setLoadingId(t.id);
    try {
      const full = await fetchTemplate(qc, t);
      if (seq !== latestPick.current) return; // superseded by a later pick
      setOpen(false);
      onPick(full);
    } catch (err) {
      if (seq !== latestPick.current) return;
      toast.error(`Could not open "${t.name}": ${errorMessage(err)}`);
    } finally {
      if (seq === latestPick.current) setLoadingId(null);
    }
  }

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button
          type="button"
          variant="outline"
          role="combobox"
          aria-expanded={open}
          aria-label={current ? `Template: ${current.name}. Change template` : 'Choose a template'}
          disabled={disabled}
          className={cn('w-full min-w-0 justify-between gap-2 font-normal sm:w-72', className)}
        >
          <LibraryBig className="text-muted-foreground" />
          <span className={cn('min-w-0 flex-1 truncate text-left', !current && 'text-muted-foreground')}>
            {current ? current.name : value && isLoading ? 'Loading…' : 'Choose a template'}
          </span>
          <ChevronsUpDown className="opacity-50" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[min(24rem,calc(100vw-2rem))] p-0" align="start">
        <Command loop>
          <CommandInput placeholder="Search templates…" />
          <CommandList className="max-h-[340px]">
            {isLoading ? (
              <div className="flex items-center justify-center gap-2 py-6 text-sm text-muted-foreground">
                <Loader2 className="size-4 animate-spin" /> Loading templates…
              </div>
            ) : isError ? (
              <div className="py-6 text-center text-sm text-destructive">Could not load templates.</div>
            ) : (
              <>
                <CommandEmpty>
                  {data?.length ? 'No templates match.' : 'No templates yet. Save one from Compose.'}
                </CommandEmpty>
                {mine.length > 0 && (
                  <CommandGroup heading="My templates">
                    {mine.map(t => (
                      <Item key={t.id} t={t} selected={t.id === value} loading={loadingId === t.id} onSelect={() => pick(t)} />
                    ))}
                  </CommandGroup>
                )}
                {shared.length > 0 && (
                  <CommandGroup heading="Shared with me">
                    {shared.map(t => (
                      <Item key={t.id} t={t} selected={t.id === value} loading={loadingId === t.id} onSelect={() => pick(t)} />
                    ))}
                  </CommandGroup>
                )}
              </>
            )}
          </CommandList>
          <CommandSeparator />
          <div className="p-1">
            <Link
              to="/templates"
              onClick={() => setOpen(false)}
              className="flex items-center gap-2 rounded-sm px-2 py-1.5 text-sm text-muted-foreground hover:bg-accent hover:text-accent-foreground focus-visible:outline-2 focus-visible:outline-[var(--ring)]"
            >
              <Settings2 className="size-4" /> Manage templates
            </Link>
          </div>
        </Command>
      </PopoverContent>
    </Popover>
  );
}
