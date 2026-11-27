// Theme tokens, ported 1:1 from the iqac-email-html skill (themes/*.json) plus the 'plain' theme.
// Keys keep the skill's snake_case so the two can be diffed side by side.
import type { ThemeId } from '../types.js';

export type HeaderStyle = 'dark' | 'light' | 'plain';
export type DetailsStyle = 'dark' | 'wash' | 'panel';

export interface ThemeColors {
  page_bg: string; card_bg: string; header_bg: string; band_bg: string;
  accent: string; accent_rule: string;
  ink: string; text: string; muted: string; rule: string; link: string;
  on_dark: string; on_dark_soft: string; on_dark_body: string; on_dark_muted: string; on_dark_faint: string;
  eyebrow_on_dark: string;
  panel_bg: string; panel_border: string; list_bg: string; dark_panel_bg: string; wash_bg: string;
  callout_bg: string; callout_border: string; callout_text: string; callout_subtext: string;
  warn_bg: string; warn_border: string; warn_text: string;
  success_bg: string; success_border: string; success_text: string;
  button_bg: string; button_text: string; button2_bg: string; button2_text: string; button2_border: string;
  bullet: string; signature_accent: string;
  /** App addition: colour of the `.kb-divider` rule (the skill's divider used `rule`). */
  divider: string;
}

export interface ThemeFontSet {
  /** Google Fonts css2 URL, or '' for none. */
  webfont_url: string;
  display: string; body: string; label: string; ui: string;
  mso_display: string; mso_body: string; mso_label: string; mso_ui: string;
}

export interface ThemeType {
  body_size: number; body_line: number;
  label_size: number; label_spacing: number; label_weight: number;
  headline_size: number; headline_weight: number;
  radius: number;
}

export interface ThemeDef {
  id: ThemeId;
  label: string;
  description: string;
  header_style: HeaderStyle;
  details_style: DetailsStyle;
  colors: ThemeColors;
  fonts: ThemeFontSet;
  type: ThemeType;
  bullet_char: string;
  /** 3-4 key colours for a theme picker. */
  swatch: string[];
}

const heritage: ThemeDef = {
  id: 'heritage',
  label: 'Heritage - Navy & Gold',
  description:
    'The original IQAC letter look: deep navy bands, gold rules, Playfair + Garamond. Formal, ceremonial. Best for offers, invitations, selections, certificates, anything addressed to an individual.',
  header_style: 'dark',
  details_style: 'dark',
  colors: {
    page_bg: '#EDEDEA', card_bg: '#FAFAF8', header_bg: '#0C1A28', band_bg: '#0A141F',
    accent: '#C5973A', accent_rule: '#C5973A',
    ink: '#0C1A28', text: '#3A3530', muted: '#7A7060', rule: '#DDD8CE', link: '#8B6914',
    on_dark: '#FFFFFF', on_dark_soft: '#E8E4DC', on_dark_body: '#C8C0B4', on_dark_muted: '#8A9AAA', on_dark_faint: '#5F6F7F',
    eyebrow_on_dark: '#C5973A',
    panel_bg: '#F3F0EA', panel_border: '#DDD8CE', list_bg: '#FDFCFB', dark_panel_bg: '#0C1A28', wash_bg: '#F3F0EA',
    callout_bg: '#FBF6EC', callout_border: '#C5973A', callout_text: '#5A4A20', callout_subtext: '#7A6A40',
    warn_bg: '#FBEDEA', warn_border: '#9B2D20', warn_text: '#6E2016',
    success_bg: '#EEF4EC', success_border: '#3F7A3A', success_text: '#28501F',
    button_bg: '#C5973A', button_text: '#0C1A28', button2_bg: '#0C1A28', button2_text: '#E8E4DC', button2_border: '#0C1A28',
    bullet: '#C5973A', signature_accent: '#8B6914',
    divider: '#C5973A',
  },
  fonts: {
    webfont_url:
      'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,600;0,700;1,600&family=EB+Garamond:ital,wght@0,400;0,500;0,600;1,400&family=Cormorant+Garamond:wght@400;600&display=swap',
    display: "'Playfair Display', Georgia, 'Times New Roman', serif",
    body: "'EB Garamond', Georgia, 'Times New Roman', serif",
    label: "'Cormorant Garamond', Georgia, 'Times New Roman', serif",
    ui: "'EB Garamond', Georgia, 'Times New Roman', serif",
    mso_display: 'Georgia, serif', mso_body: 'Georgia, serif', mso_label: 'Georgia, serif', mso_ui: 'Georgia, serif',
  },
  type: { body_size: 16.5, body_line: 1.75, label_size: 11, label_spacing: 3, label_weight: 600, headline_size: 26, headline_weight: 700, radius: 2 },
  bullet_char: '▸',
  swatch: ['#0C1A28', '#C5973A', '#FAFAF8', '#8B6914'],
};

const meridian: ThemeDef = {
  id: 'meridian',
  label: 'Meridian - Spruce & Signal Green',
  description:
    'The IQAC website palette: near-black spruce bands, signal-green accents, mono labels. Contemporary and crisp. Best for system-style notices, reminders, status updates, data-heavy mails, and anything that should match the website.',
  header_style: 'dark',
  details_style: 'wash',
  colors: {
    page_bg: '#ECEDE6', card_bg: '#FFFFFF', header_bg: '#0B2016', band_bg: '#081710',
    accent: '#17D059', accent_rule: '#17D059',
    ink: '#0B1410', text: '#2C3631', muted: '#5C655F', rule: '#D9DBD2', link: '#0E7A37',
    on_dark: '#EAF0EC', on_dark_soft: '#B9C7BE', on_dark_body: '#B9C7BE', on_dark_muted: '#8FA097', on_dark_faint: '#6B7C73',
    eyebrow_on_dark: '#17D059',
    panel_bg: '#F7F6F1', panel_border: '#D9DBD2', list_bg: '#FFFFFF', dark_panel_bg: '#0B2016', wash_bg: '#E4F7EA',
    callout_bg: '#F7F6F1', callout_border: '#17D059', callout_text: '#2C3631', callout_subtext: '#5C655F',
    warn_bg: '#F9ECEA', warn_border: '#9B2D20', warn_text: '#9B2D20',
    success_bg: '#E4F7EA', success_border: '#17D059', success_text: '#0B4A22',
    button_bg: '#17D059', button_text: '#0B1410', button2_bg: '#0B2016', button2_text: '#EAF0EC', button2_border: '#0B2016',
    bullet: '#17D059', signature_accent: '#0E7A37',
    divider: '#D9DBD2',
  },
  fonts: {
    webfont_url:
      'https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600;700&family=EB+Garamond:ital,wght@0,400;0,500;0,600;1,400&display=swap',
    display: "'Playfair Display', Georgia, 'Times New Roman', serif",
    body: "'EB Garamond', Georgia, 'Times New Roman', serif",
    label: "'SFMono-Regular', Consolas, 'Liberation Mono', monospace",
    ui: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif",
    mso_display: 'Georgia, serif', mso_body: 'Georgia, serif', mso_label: 'Consolas, monospace', mso_ui: "'Segoe UI', Arial, sans-serif",
  },
  type: { body_size: 17, body_line: 1.7, label_size: 11, label_spacing: 1.6, label_weight: 400, headline_size: 26, headline_weight: 600, radius: 2 },
  bullet_char: '▸',
  swatch: ['#0B2016', '#17D059', '#FFFFFF', '#0E7A37'],
};

const letterhead: ThemeDef = {
  id: 'letterhead',
  label: 'Letterhead - Ivory & Oxblood',
  description:
    'A light, printed-letterhead look: no dark bands, ivory paper, oxblood double rule, centred crest. Quiet and official. Best for circulars, formal notices to faculty/deans, minutes, apologies, and mail likely to be printed or forwarded up the chain.',
  header_style: 'light',
  details_style: 'panel',
  colors: {
    page_bg: '#EFEDE7', card_bg: '#FFFEFB', header_bg: '#FFFEFB', band_bg: '#FFFEFB',
    accent: '#7A1F2B', accent_rule: '#7A1F2B',
    ink: '#1E1A17', text: '#3B3632', muted: '#77706A', rule: '#E2DDD4', link: '#7A1F2B',
    on_dark: '#1E1A17', on_dark_soft: '#3B3632', on_dark_body: '#3B3632', on_dark_muted: '#77706A', on_dark_faint: '#958E87',
    eyebrow_on_dark: '#7A1F2B',
    panel_bg: '#F7F4EE', panel_border: '#E2DDD4', list_bg: '#FFFEFB', dark_panel_bg: '#2A1F1D', wash_bg: '#F7F4EE',
    callout_bg: '#F8F1EF', callout_border: '#7A1F2B', callout_text: '#4A1D22', callout_subtext: '#77706A',
    warn_bg: '#F8EDEA', warn_border: '#9B2D20', warn_text: '#7A2216',
    success_bg: '#EFF3EC', success_border: '#4A6B3D', success_text: '#2F4626',
    button_bg: '#7A1F2B', button_text: '#FFFFFF', button2_bg: '#FFFEFB', button2_text: '#7A1F2B', button2_border: '#7A1F2B',
    bullet: '#7A1F2B', signature_accent: '#7A1F2B',
    divider: '#7A1F2B',
  },
  fonts: {
    webfont_url:
      'https://fonts.googleapis.com/css2?family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&family=EB+Garamond:ital,wght@0,400;0,500;0,600;1,400&display=swap',
    display: "'Libre Baskerville', Baskerville, Georgia, 'Times New Roman', serif",
    body: "'EB Garamond', Georgia, 'Times New Roman', serif",
    label: "'EB Garamond', Georgia, 'Times New Roman', serif",
    ui: "'EB Garamond', Georgia, 'Times New Roman', serif",
    mso_display: 'Georgia, serif', mso_body: 'Georgia, serif', mso_label: 'Georgia, serif', mso_ui: 'Georgia, serif',
  },
  type: { body_size: 17, body_line: 1.7, label_size: 12, label_spacing: 2.4, label_weight: 600, headline_size: 23, headline_weight: 700, radius: 0 },
  bullet_char: '▸',
  swatch: ['#7A1F2B', '#FFFEFB', '#1E1A17', '#EFEDE7'],
};

const SYSTEM_SANS = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif";

const slate: ThemeDef = {
  id: 'slate',
  label: 'Slate - Modern Sans',
  description:
    'Clean, modern sans-serif: charcoal header, cobalt accent, system fonts that render identically everywhere (no webfont dependency). Best for student-facing mail, event logistics, workshops, quick updates, and anything read mostly on phones.',
  header_style: 'dark',
  details_style: 'wash',
  colors: {
    page_bg: '#EEF0F3', card_bg: '#FFFFFF', header_bg: '#111827', band_bg: '#0B1220',
    accent: '#3B6CF6', accent_rule: '#3B6CF6',
    ink: '#111827', text: '#374151', muted: '#6B7280', rule: '#E5E7EB', link: '#2550D6',
    on_dark: '#F9FAFB', on_dark_soft: '#D1D5DB', on_dark_body: '#D1D5DB', on_dark_muted: '#9CA3AF', on_dark_faint: '#6B7280',
    eyebrow_on_dark: '#8FB0FF',
    panel_bg: '#F3F4F6', panel_border: '#E5E7EB', list_bg: '#FFFFFF', dark_panel_bg: '#111827', wash_bg: '#EEF3FF',
    callout_bg: '#EEF3FF', callout_border: '#3B6CF6', callout_text: '#1E2A4A', callout_subtext: '#4B5563',
    warn_bg: '#FEF2F2', warn_border: '#DC2626', warn_text: '#991B1B',
    success_bg: '#ECFDF3', success_border: '#16A34A', success_text: '#14532D',
    button_bg: '#2550D6', button_text: '#FFFFFF', button2_bg: '#111827', button2_text: '#F9FAFB', button2_border: '#111827',
    bullet: '#3B6CF6', signature_accent: '#2550D6',
    divider: '#E5E7EB',
  },
  fonts: {
    webfont_url: '',
    display: SYSTEM_SANS, body: SYSTEM_SANS, label: SYSTEM_SANS, ui: SYSTEM_SANS,
    mso_display: "'Segoe UI', Arial, sans-serif", mso_body: "'Segoe UI', Arial, sans-serif",
    mso_label: "'Segoe UI', Arial, sans-serif", mso_ui: "'Segoe UI', Arial, sans-serif",
  },
  type: { body_size: 15.5, body_line: 1.65, label_size: 11, label_spacing: 1.2, label_weight: 600, headline_size: 24, headline_weight: 700, radius: 6 },
  bullet_char: '•',
  swatch: ['#111827', '#3B6CF6', '#FFFFFF', '#EEF0F3'],
};

// App addition (not in the skill): a formal, minimal letter. White card, thin KIIT-navy top rule,
// text-only header, Georgia headings over Arial body, no web fonts.
const GEORGIA = "Georgia, 'Times New Roman', Times, serif";
const ARIAL = 'Arial, Helvetica, sans-serif';
const plain: ThemeDef = {
  id: 'plain',
  label: 'Plain - Formal Letter',
  description:
    'A minimal formal letter: white page, a thin KIIT-navy rule, text-only header, Georgia and Arial. No banners or web fonts. Best for routine correspondence, replies, and mail that should read like a personal letter.',
  header_style: 'plain',
  details_style: 'panel',
  colors: {
    page_bg: '#F2F2EF', card_bg: '#FFFFFF', header_bg: '#FFFFFF', band_bg: '#FFFFFF',
    accent: '#13294B', accent_rule: '#13294B',
    ink: '#1A1A1A', text: '#333333', muted: '#6B6B6B', rule: '#E2E2E2', link: '#1A4D8F',
    on_dark: '#1A1A1A', on_dark_soft: '#333333', on_dark_body: '#333333', on_dark_muted: '#6B6B6B', on_dark_faint: '#8A8A8A',
    eyebrow_on_dark: '#13294B',
    panel_bg: '#F6F6F4', panel_border: '#E2E2E2', list_bg: '#FFFFFF', dark_panel_bg: '#13294B', wash_bg: '#F3F5F9',
    callout_bg: '#F3F5F9', callout_border: '#13294B', callout_text: '#1F2A3A', callout_subtext: '#555555',
    warn_bg: '#FDF1EF', warn_border: '#A8322A', warn_text: '#7A2318',
    success_bg: '#EFF6EF', success_border: '#3A7A3A', success_text: '#234D23',
    button_bg: '#13294B', button_text: '#FFFFFF', button2_bg: '#FFFFFF', button2_text: '#13294B', button2_border: '#13294B',
    bullet: '#13294B', signature_accent: '#13294B',
    divider: '#E2E2E2',
  },
  fonts: {
    webfont_url: '',
    display: GEORGIA, body: ARIAL, label: ARIAL, ui: ARIAL,
    mso_display: 'Georgia, serif', mso_body: 'Arial, sans-serif', mso_label: 'Arial, sans-serif', mso_ui: 'Arial, sans-serif',
  },
  type: { body_size: 15, body_line: 1.6, label_size: 11, label_spacing: 1.2, label_weight: 600, headline_size: 22, headline_weight: 700, radius: 2 },
  bullet_char: '•',
  swatch: ['#13294B', '#FFFFFF', '#333333', '#F2F2EF'],
};

export const THEME_DEFS: Record<ThemeId, ThemeDef> = { heritage, meridian, letterhead, slate, plain };
export const THEME_ORDER: ThemeId[] = ['heritage', 'meridian', 'letterhead', 'slate', 'plain'];
export const DEFAULT_THEME: ThemeId = 'heritage';

/** Theme definition by id; unknown ids fall back to heritage (the skill's default). */
export function getTheme(id: ThemeId | string | null | undefined): ThemeDef {
  return (id && (THEME_DEFS as Record<string, ThemeDef>)[id]) || THEME_DEFS[DEFAULT_THEME];
}

/** Python's round(): halves go to the even neighbour, so blends match build_email.py exactly. */
function pyRound(v: number): number {
  const f = Math.floor(v);
  if (Math.abs(v - f - 0.5) < 1e-9) return f % 2 === 0 ? f : f + 1;
  return Math.round(v);
}

/** Blend hex colour a over b at opacity t (0..1). Stands in for rgba(), which Outlook ignores. */
export function mix(a: string, b: string, t: number): string {
  const pa = a.replace('#', ''), pb = b.replace('#', '');
  let out = '#';
  for (const i of [0, 2, 4]) {
    const x = parseInt(pa.slice(i, i + 2), 16), y = parseInt(pb.slice(i, i + 2), 16);
    out += pyRound(x * t + y * (1 - t)).toString(16).padStart(2, '0').toUpperCase();
  }
  return out;
}
