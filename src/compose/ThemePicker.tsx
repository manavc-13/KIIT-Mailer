// Letterhead theme cards (rich mode only). Uses shared THEMES; falls back to a built-in list
// while the themes module is still a stub.
import { useRef, type KeyboardEvent } from 'react';
import { Check } from 'lucide-react';
import { cn } from '@/lib/utils';
import { THEMES, type ThemeMeta } from '../../shared/themes';
import type { ThemeId } from '../../shared/types';
import { useComposeStore } from './store';
import { useSelectedSender } from './model';

const FALLBACK_THEMES: ThemeMeta[] = [
  { id: 'heritage', label: 'Heritage', description: 'Navy and gold, formal letterhead', swatch: ['#0c1a28', '#b8892f', '#f6f5f1'] },
  { id: 'meridian', label: 'Meridian', description: 'Spruce and green', swatch: ['#1f4d45', '#3f8f5a', '#f3f6f2'] },
  { id: 'letterhead', label: 'Letterhead', description: 'Ivory and oxblood', swatch: ['#6b1d1d', '#f7f1e3', '#2b2b2b'] },
  { id: 'slate', label: 'Slate', description: 'Modern sans-serif', swatch: ['#334155', '#64748b', '#f1f5f9'] },
  { id: 'plain', label: 'Plain', description: 'No letterhead', swatch: ['#ffffff', '#d4d4d4', '#171717'] },
];

export function useThemes(): ThemeMeta[] {
  return THEMES.length ? THEMES : FALLBACK_THEMES;
}

export function themeLabel(id: ThemeId, themes: ThemeMeta[]): string {
  return themes.find(t => t.id === id)?.label ?? id;
}

export function ThemePicker() {
  const themes = useThemes();
  const theme = useComposeStore(s => s.theme);
  const edit = useComposeStore(s => s.edit);
  const { sender } = useSelectedSender();
  const refs = useRef<(HTMLButtonElement | null)[]>([]);

  const choose = (id: ThemeId) => edit({ theme: id, themeTouched: true });
  const onKey = (e: KeyboardEvent, i: number) => {
    const dir = e.key === 'ArrowRight' || e.key === 'ArrowDown' ? 1 : e.key === 'ArrowLeft' || e.key === 'ArrowUp' ? -1 : 0;
    if (!dir) return;
    e.preventDefault();
    const n = (i + dir + themes.length) % themes.length;
    choose(themes[n].id);
    refs.current[n]?.focus();
  };

  return (
    <div className="mb-4">
      <div className="mb-2 flex flex-wrap items-baseline justify-between gap-2">
        <span id="theme-label" className="text-sm font-medium">Letterhead theme</span>
        {sender && (
          <span className="text-xs text-muted-foreground">
            {sender.displayName}'s default: {themeLabel(sender.defaultTheme, themes)}
            {theme !== sender.defaultTheme && (
              <button type="button" className="ml-2 underline underline-offset-2 hover:text-foreground" onClick={() => choose(sender.defaultTheme)}>
                Use default
              </button>
            )}
          </span>
        )}
      </div>
      <div role="radiogroup" aria-labelledby="theme-label" className="grid grid-cols-2 gap-2 sm:grid-cols-3 xl:grid-cols-5">
        {themes.map((t, i) => {
          const active = t.id === theme;
          return (
            <button
              key={t.id}
              ref={el => { refs.current[i] = el; }}
              type="button" role="radio" aria-checked={active} tabIndex={active ? 0 : -1}
              onClick={() => choose(t.id)} onKeyDown={e => onKey(e, i)}
              className={cn(
                'relative flex flex-col gap-2 rounded-md border p-2.5 text-left transition-colors outline-none focus-visible:ring-2 focus-visible:ring-ring',
                active ? 'border-primary ring-1 ring-primary dark:border-gold dark:ring-gold' : 'hover:bg-muted/60',
              )}
            >
              <span className="flex h-7 overflow-hidden rounded-sm ring-1 ring-border" aria-hidden>
                {t.swatch.slice(0, 4).map((c, k) => <span key={k} className="flex-1" style={{ background: c }} />)}
              </span>
              <span className="text-[13px] font-medium leading-tight">{t.label}</span>
              <span className="line-clamp-2 text-[11px] leading-snug text-muted-foreground">{t.description}</span>
              {active && <Check className="absolute right-2 top-2 size-3.5 rounded-full bg-primary p-0.5 text-primary-foreground" />}
            </button>
          );
        })}
      </div>
    </div>
  );
}
