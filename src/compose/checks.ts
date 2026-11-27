// Deliverability / readiness checklist for the Review step: server lint (RenderResult.lint)
// plus local checks (recipients, attachments, quota, test send, placeholders).
import { formatBytes, placeholdersIn } from '../../shared/text';
import type { BuildJobRecipientsResult } from '../../shared/recipients';
import type { ComposeContent, JobMode, RenderResult, SenderAccount } from '../../shared/types';
import type { AttachmentMeter } from './attachmentMath';
import type { ComposeStep } from './store';

export type CheckStatus = 'pass' | 'warn' | 'fail' | 'info';
export interface Check {
  id: string;
  status: CheckStatus;
  title: string;
  detail?: string;
  /** Step where the user can fix it. */
  fixStep?: ComposeStep;
  /** Local failures block sending; lint failures need an explicit "send anyway". */
  blocking?: boolean;
}

export interface CheckInput {
  sender: SenderAccount | null;
  content: ComposeContent;
  mode: JobMode;
  build: BuildJobRecipientsResult;
  placeholders: string[];
  attachmentCount: number;
  attachmentBlocking: string | null;
  meter: AttachmentMeter;
  testSent: boolean;
  render: RenderResult | null;
  renderPending: boolean;
}

const LINT_STATUS: Record<'error' | 'warn' | 'info', CheckStatus> = { error: 'fail', warn: 'warn', info: 'info' };

export function quotaAfter(sender: SenderAccount | null, cost: number): { remaining: number; after: number } {
  const remaining = sender ? Math.max(0, sender.dailyLimit - sender.usedToday) : 0;
  return { remaining, after: remaining - cost };
}

export function buildChecks(i: CheckInput): Check[] {
  const out: Check[] = [];
  const { build } = i;

  // Sender
  if (!i.sender) out.push({ id: 'sender', status: 'fail', blocking: true, title: 'Choose a From account', detail: 'Pick a sender account at the top of the page.' });

  // Recipients
  const skipped = build.skippedInvalid.length + build.skippedMissing;
  if (!build.recipientCount) {
    out.push({ id: 'recipients', status: 'fail', blocking: true, fixStep: 'recipients', title: 'No valid recipients', detail: 'Add at least one valid email address.' });
  } else if (skipped) {
    out.push({
      id: 'recipients', status: 'warn', fixStep: 'recipients',
      title: `${skipped.toLocaleString()} row${skipped === 1 ? '' : 's'} will be skipped`,
      detail: [
        build.skippedInvalid.length ? `${build.skippedInvalid.length} invalid address${build.skippedInvalid.length === 1 ? '' : 'es'} (e.g. ${build.skippedInvalid.slice(0, 3).join(', ')})` : '',
        build.skippedMissing ? `${build.skippedMissing} without an email` : '',
      ].filter(Boolean).join('; '),
    });
  } else {
    out.push({ id: 'recipients', status: 'pass', title: `${build.recipientCount.toLocaleString()} valid recipient${build.recipientCount === 1 ? '' : 's'}` });
  }
  if (build.invalidCopies.length) {
    out.push({ id: 'copies', status: 'warn', fixStep: 'recipients', title: 'Some CC/BCC addresses are invalid and will be dropped', detail: build.invalidCopies.slice(0, 5).join(', ') });
  }

  // Content
  if (!i.content.subject.trim()) out.push({ id: 'subject', status: 'fail', blocking: true, fixStep: 'compose', title: 'Subject is empty' });
  const bodyText = i.content.bodyHtml.replace(/<style[\s\S]*?<\/style>/gi, '').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
  if (!bodyText && !/<img\b/i.test(i.content.bodyHtml)) out.push({ id: 'body', status: 'fail', blocking: true, fixStep: 'compose', title: 'Message body is empty' });

  // Placeholders
  const used = placeholdersIn(i.content.subject, i.content.preheader, i.content.bodyHtml, i.content.bodyText);
  if (used.length) {
    if (i.mode === 'bcc_broadcast') {
      out.push({
        id: 'placeholders', status: 'warn', fixStep: 'compose',
        title: 'Placeholders are not replaced in a BCC broadcast',
        detail: `${used.map(p => `{${p}}`).join(', ')} will appear as written. Remove them or switch to Personalized.`,
      });
    } else {
      const known = new Set(i.placeholders.map(p => p.toLowerCase()));
      const unknown = used.filter(p => !known.has(p.toLowerCase()));
      if (unknown.length) {
        out.push({
          id: 'placeholders', status: 'warn', fixStep: 'recipients',
          title: 'Some placeholders have no matching column',
          detail: `${unknown.map(p => `{${p}}`).join(', ')} will appear as written.`,
        });
      } else {
        out.push({ id: 'placeholders', status: 'pass', title: 'All placeholders match a recipient column' });
      }
    }
  }

  // Attachments
  if (i.attachmentBlocking) {
    out.push({ id: 'attachments', status: 'fail', blocking: true, fixStep: 'compose', title: 'Attachments not ready', detail: i.attachmentBlocking });
  } else if (i.meter.level === 'danger') {
    out.push({ id: 'attachments', status: 'fail', blocking: true, fixStep: 'compose', title: 'Message is over Gmail\'s size limit', detail: i.meter.label });
  } else if (i.meter.largeWarning) {
    out.push({ id: 'attachments', status: 'warn', fixStep: 'compose', title: `Large attachments (${formatBytes(i.meter.rawBytes)})`, detail: 'Some mail servers reject messages this size. Consider a Drive link.' });
  } else if (i.attachmentCount) {
    out.push({ id: 'attachments', status: 'pass', title: `Attachments within limits (${formatBytes(i.meter.rawBytes)})` });
  }

  // Quota
  if (i.sender) {
    const { remaining, after } = quotaAfter(i.sender, build.quotaCost);
    if (after < 0) {
      out.push({
        id: 'quota', status: 'fail', blocking: true,
        title: 'Not enough daily quota',
        detail: `This send counts ${build.quotaCost.toLocaleString()} recipients (To + CC + BCC) but ${i.sender.email} has ${remaining.toLocaleString()} left today. Reduce the list or send the rest tomorrow.`,
      });
    } else if (after < i.sender.dailyLimit * 0.1) {
      out.push({ id: 'quota', status: 'warn', title: `Quota nearly used: ${after.toLocaleString()} left after this send` });
    } else {
      out.push({ id: 'quota', status: 'pass', title: `Within daily quota (${after.toLocaleString()} left after this send)` });
    }
  }

  // Test send
  out.push(i.testSent
    ? { id: 'test', status: 'pass', title: 'Test email sent this session' }
    : { id: 'test', status: 'warn', title: 'No test email sent yet', detail: 'Send yourself a test and check it on your phone before sending to everyone.' });

  // Server lint
  if (i.render) {
    const lint = i.render.lint;
    if (!lint.some(l => l.level !== 'info')) out.push({ id: 'lint', status: 'pass', title: 'Content check passed', detail: `Rendered size ${formatBytes(i.render.sizeBytes)}.` });
    lint.forEach((l, k) => out.push({ id: `lint-${l.code}-${k}`, status: LINT_STATUS[l.level], fixStep: 'compose', title: l.message }));
    if (i.render.webFontsUsed.length) {
      out.push({ id: 'fonts', status: 'info', title: `Web fonts fall back in Gmail and Outlook: ${i.render.webFontsUsed.join(', ')}` });
    }
  } else {
    out.push({ id: 'lint', status: 'info', title: i.renderPending ? 'Checking content…' : 'Content check runs once a sender is chosen' });
  }

  out.push({ id: 'suppressed', status: 'info', title: 'Addresses that unsubscribed or bounced are skipped automatically' });
  return out;
}

const ORDER: Record<CheckStatus, number> = { fail: 0, warn: 1, pass: 2, info: 3 };
export function sortChecks(checks: Check[]): Check[] {
  return [...checks].sort((a, b) => ORDER[a.status] - ORDER[b.status]);
}
