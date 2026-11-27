// Brand defaults (from the skill's assets/brand.json) and the starter signature.
// Personal signature presets in brand.json are intentionally not ported: signatures are per
// sender account (`signatureHtml`), and starterSignatureHtml() bootstraps a new one.
import { escapeHtml } from '../text.js';
import type { Brand, SignatureInput } from './model.js';

export const DEFAULT_BRAND: Brand = {
  logoUrl: 'https://pub-5b9367aff065489ea09a6296c3ef1745.r2.dev/brand/kiit-logo.png',
  logoAlt: 'KIIT',
  logoWidth: 88,
  eyebrow: 'Internal Quality Assurance Cell',
  title: 'Kalinga Institute of Industrial Technology',
  tagline: '',
  footerLine: 'Internal Quality Assurance Cell (IQAC) · KIIT Deemed to be University',
  disclaimer: 'This message is intended solely for the named recipient. If received in error, please disregard.',
};

/** DEFAULT_BRAND with overrides applied. `undefined`/`null` keep the default; '' clears a field. */
export function resolveBrand(over?: Partial<Brand> | null): Brand {
  const out: Brand = { ...DEFAULT_BRAND };
  if (!over) return out;
  for (const k of Object.keys(over) as (keyof Brand)[]) {
    const v = over[k];
    if (v === undefined || v === null) continue;
    (out as unknown as Record<string, unknown>)[k] = v;
  }
  if (!Number.isFinite(out.logoWidth) || out.logoWidth <= 0) out.logoWidth = DEFAULT_BRAND.logoWidth;
  return out;
}

export const SIGNATURE_DEFAULT_DEPARTMENT = 'Internal Quality Assurance Cell (IQAC)';
export const SIGNATURE_ORG_LINE = 'KIIT Deemed to be University';

/**
 * Email-safe starter signature for a new sender account. Theme-neutral: no inline colours or
 * fonts, only classes that themeCss() styles inside `.kb-signature` (the shell wraps the
 * signature in that cell). If an editor drops the classes, the plain <p>/<strong>/<a> markup
 * still renders as a tidy signature.
 *
 *   Warm regards,
 *   **Name**
 *   Title
 *   Internal Quality Assurance Cell (IQAC)
 *   KIIT Deemed to be University
 *   E: email | M: phone
 */
export function starterSignatureHtml(input: SignatureInput): string {
  const closing = (input.closing ?? '').trim() || 'Warm regards,';
  const name = (input.name ?? '').trim();
  const title = (input.title ?? '').trim();
  const dept = (input.department ?? '').trim() || SIGNATURE_DEFAULT_DEPARTMENT;
  const email = (input.email ?? '').trim();
  const phone = (input.phone ?? '').trim();

  const parts: string[] = [];
  parts.push(`<p class="kb-signature-closing">${escapeHtml(closing)}</p>`);
  if (name) parts.push(`<p class="kb-signature-name"><strong>${escapeHtml(name)}</strong></p>`);
  if (title) parts.push(`<p class="kb-signature-title">${escapeHtml(title)}</p>`);
  parts.push(`<p class="kb-signature-org">${escapeHtml(dept)}<br>${escapeHtml(SIGNATURE_ORG_LINE)}</p>`);

  const contacts: string[] = [];
  if (email) {
    contacts.push(`<strong>E:</strong> <a href="mailto:${escapeHtml(email)}">${escapeHtml(email)}</a>`);
  }
  if (phone) {
    const tel = phone.replace(/[^\d+]/g, '');
    contacts.push(
      tel ? `<strong>M:</strong> <a href="tel:${escapeHtml(tel)}">${escapeHtml(phone)}</a>` : `<strong>M:</strong> ${escapeHtml(phone)}`,
    );
  }
  if (contacts.length) {
    parts.push(`<p class="kb-signature-contact">${contacts.join('&nbsp;<span class="kb-signature-sep">|</span>&nbsp;')}</p>`);
  }
  return parts.join('');
}
