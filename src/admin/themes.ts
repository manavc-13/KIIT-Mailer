// Theme choices for the sender account "default theme" picker. Uses the shared
// theme catalog when it is populated, otherwise a fixed fallback list.
import { THEMES } from '../../shared/themes';
import type { ThemeId } from '../../shared/types';

export interface ThemeOption { id: ThemeId; label: string; description: string; swatch: string[] }

const FALLBACK: ThemeOption[] = [
  { id: 'heritage', label: 'Heritage', description: 'Navy and gold, formal letterhead', swatch: ['#0c1a28', '#b8892f'] },
  { id: 'meridian', label: 'Meridian', description: 'Spruce and green', swatch: ['#1f4d45', '#5c9a6b'] },
  { id: 'letterhead', label: 'Letterhead', description: 'Ivory and oxblood', swatch: ['#f6f1e7', '#7a1f2b'] },
  { id: 'slate', label: 'Slate', description: 'Modern sans serif', swatch: ['#2f3a45', '#6b7a89'] },
  { id: 'plain', label: 'Plain', description: 'No letterhead', swatch: ['#ffffff', '#9aa3ad'] },
];

export function themeOptions(): ThemeOption[] {
  return THEMES.length ? THEMES.map(t => ({ id: t.id, label: t.label, description: t.description, swatch: t.swatch })) : FALLBACK;
}

export function themeLabel(id: ThemeId | string): string {
  return themeOptions().find(t => t.id === id)?.label ?? id;
}
