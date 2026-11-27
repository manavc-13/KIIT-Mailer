// Name + category inputs shared by the save and rename dialogs.
import { useId } from 'react';
import { AlertCircle } from 'lucide-react';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { cn } from '@/lib/utils';

export function NameCategoryFields({
  name, onName, category, onCategory, suggestions, autoFocus = true, nameError,
}: {
  name: string; onName: (v: string) => void;
  category: string; onCategory: (v: string) => void;
  suggestions: string[]; autoFocus?: boolean; nameError?: string | null;
}) {
  const id = useId();
  const listId = `${id}-categories`;
  const errId = `${id}-name-error`;
  const filtered = suggestions.filter(s => s.toLowerCase() !== category.trim().toLowerCase());
  return (
    <div className="grid gap-4">
      <div className="grid gap-1.5">
        <Label htmlFor={`${id}-name`}>Template name</Label>
        <Input
          id={`${id}-name`}
          value={name}
          maxLength={120}
          autoFocus={autoFocus}
          onChange={e => onName(e.target.value)}
          placeholder="e.g. NAAC peer-team visit circular"
          aria-invalid={!!nameError}
          aria-describedby={nameError ? errId : undefined}
          required
        />
        {nameError && (
          <p id={errId} className="flex items-center gap-1 text-xs text-destructive"><AlertCircle className="size-3.5" />{nameError}</p>
        )}
      </div>
      <div className="grid gap-1.5">
        <Label htmlFor={`${id}-category`}>Category <span className="font-normal text-muted-foreground">(optional)</span></Label>
        <Input
          id={`${id}-category`}
          value={category}
          maxLength={60}
          list={listId}
          onChange={e => onCategory(e.target.value)}
          placeholder="e.g. Circulars, Invitations, Reminders"
        />
        <datalist id={listId}>
          {suggestions.map(s => <option key={s} value={s} />)}
        </datalist>
        {filtered.length > 0 && (
          <div className="flex flex-wrap gap-1.5" aria-label="Existing categories">
            {filtered.slice(0, 8).map(s => (
              <button
                key={s}
                type="button"
                onClick={() => onCategory(s)}
                className={cn(
                  'rounded-full border bg-background px-2.5 py-0.5 text-xs text-muted-foreground transition-colors',
                  'hover:border-[var(--gold)] hover:text-foreground focus-visible:outline-2 focus-visible:outline-[var(--ring)]',
                )}
              >
                {s}
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
