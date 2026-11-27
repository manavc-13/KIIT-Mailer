// Public shapes of the themes module (re-exported from ./index.ts).
import type { ThemeId } from '../types.js';

export interface ThemeMeta { id: ThemeId; label: string; description: string; swatch: string[] }

export interface Brand {
  logoUrl: string; logoAlt: string; logoWidth: number;
  eyebrow: string; title: string; tagline: string;
  footerLine: string; disclaimer: string;
}

export interface ShellOptions {
  theme: ThemeId;
  /** Email-safe body fragment (editor serializer output, placeholders already substituted). */
  bodyHtml: string;
  /** Sender signature fragment, or null to omit. */
  signatureHtml: string | null;
  subject: string;
  /** One-click unsubscribe URL; when set the footer shows an "Unsubscribe" link. */
  unsubscribeUrl?: string | null;
  brand?: Partial<Brand>;
  /** Optional big title in the header band (the skill's `headline`). Omit for a plain-letter feel. */
  headline?: string | null;
  /** Optional small-caps line under the headline (the skill's `subheadline`). */
  subheadline?: string | null;
}

export interface SignatureInput {
  name: string;
  title?: string | null;
  /** Replaces the default "Internal Quality Assurance Cell (IQAC)" line. */
  department?: string | null;
  email: string;
  phone?: string | null;
  /** Defaults to "Warm regards,". */
  closing?: string | null;
}
