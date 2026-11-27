// Shared constants for the email editor (pure, no DOM).

/** Matches a placeholder token like {Name} or {Roll No}. Must start with a letter. */
export const FIELD_SOURCE = String.raw`\{([A-Za-z][A-Za-z0-9 _.\-/]{0,48})\}`;
export const fieldRegex = (flags = 'g') => new RegExp(FIELD_SOURCE, flags);

/** Pixel width of one indent level (paragraph margin-left). */
export const INDENT_PX = 32;
export const MAX_INDENT = 8;

/** Email content width: images are capped at this. */
export const MAX_IMAGE_WIDTH = 600;

/** Border colour used for data tables in the sent mail. */
export const TABLE_BORDER = '#c9ced6';

export interface Swatch { label: string; value: string }

/** Formal text-colour palette (navy / gold / oxblood / spruce / greys). */
export const TEXT_COLORS: Swatch[] = [
  { label: 'Navy', value: '#0c1a28' },
  { label: 'Royal navy', value: '#1f3864' },
  { label: 'KIIT blue', value: '#1d4e89' },
  { label: 'Gold', value: '#b8892f' },
  { label: 'Antique gold', value: '#8a6420' },
  { label: 'Oxblood', value: '#7a1f2b' },
  { label: 'Crimson', value: '#a4262c' },
  { label: 'Spruce', value: '#1f5e4a' },
  { label: 'Forest', value: '#2f6b2a' },
  { label: 'Slate', value: '#4a5868' },
  { label: 'Charcoal', value: '#333333' },
  { label: 'Grey', value: '#6b7280' },
  { label: 'Silver', value: '#a0a7b1' },
  { label: 'Black', value: '#000000' },
  { label: 'White', value: '#ffffff' },
];

/** Soft highlight colours that stay readable with dark text. */
export const HIGHLIGHT_COLORS: Swatch[] = [
  { label: 'Pale gold', value: '#f6efdf' },
  { label: 'Pale yellow', value: '#fff3bf' },
  { label: 'Ivory', value: '#f8f6f0' },
  { label: 'Pale blue', value: '#e6eef8' },
  { label: 'Pale green', value: '#e7f3ea' },
  { label: 'Pale rose', value: '#f9e6e8' },
  { label: 'Pale grey', value: '#eef0f3' },
  { label: 'Navy', value: '#0c1a28' },
  { label: 'Gold', value: '#b8892f' },
];

export type CalloutVariant = 'info' | 'warn' | 'success';
export const CALLOUT_VARIANTS: { id: CalloutVariant; label: string }[] = [
  { id: 'info', label: 'Information' },
  { id: 'warn', label: 'Warning' },
  { id: 'success', label: 'Success' },
];

export type ButtonVariant = 'primary' | 'secondary';
export type ButtonAlign = 'left' | 'center';
