// Attachment size budget, ported from legacy/public/js/attachments.js.
//
// Gmail caps the total ENCODED message size at 25 MB; base64 inflates raw bytes by ~4/3, so the
// usable raw budget is roughly 18 MB. We size against the encoded number (attachments + body +
// MIME overhead), which is what the server enforces too. A soft warning shows above 10 MB because
// many external mail servers reject messages smaller than Gmail's own limit.
import { encodedSize, formatBytes } from '../../shared/text';
import type { UploadConfig } from '../../shared/types';

export type BudgetConfig = Pick<UploadConfig, 'maxRawAttachmentBytes' | 'maxEncodedMessageBytes' | 'softWarnBytes' | 'mimeOverheadBytes'>;

/** Conservative defaults used until /api/uploads/config answers. */
export const DEFAULT_BUDGET: BudgetConfig = {
  maxRawAttachmentBytes: 18 * 1024 * 1024,
  maxEncodedMessageBytes: 25 * 1024 * 1024,
  softWarnBytes: 10 * 1024 * 1024,
  mimeOverheadBytes: 96 * 1024,
};

export function utf8Bytes(...parts: (string | null | undefined)[]): number {
  const enc = new TextEncoder();
  return parts.reduce((n, p) => n + (p ? enc.encode(p).length : 0), 0);
}

export interface AttachmentMeter {
  rawBytes: number;
  encodedBytes: number;
  /** 0-100, clamped. */
  percent: number;
  level: 'ok' | 'warn' | 'danger';
  /** Raw attachments exceed the soft-warn size (suggest a Drive link). */
  largeWarning: boolean;
  label: string;
}

export function computeMeter(rawBytes: number, bodyBytes: number, cfg: BudgetConfig = DEFAULT_BUDGET): AttachmentMeter {
  const encodedBytes = encodedSize(rawBytes) + bodyBytes + cfg.mimeOverheadBytes;
  const percent = Math.min(100, Math.round((encodedBytes / cfg.maxEncodedMessageBytes) * 100));
  const level = encodedBytes > cfg.maxEncodedMessageBytes ? 'danger' : encodedBytes > cfg.softWarnBytes ? 'warn' : 'ok';
  return {
    rawBytes, encodedBytes, percent, level,
    largeWarning: rawBytes > cfg.softWarnBytes,
    label: `${formatBytes(rawBytes)} raw ≈ ${formatBytes(encodedBytes)} encoded / ${formatBytes(cfg.maxEncodedMessageBytes)} budget (${percent}%)`,
  };
}

export interface FileLike { name: string; size: number }
export interface AdmitResult<F extends FileLike> { accepted: F[]; rejected: { file: F; reason: string }[] }

/**
 * Decide which newly picked files fit. Empty files (or folders) are skipped, files over the
 * per-file raw cap are rejected, and files that would push the encoded message over the
 * Gmail budget are rejected (earlier files in the same batch count towards later ones).
 */
export function admitFiles<F extends FileLike>(
  files: F[], existingRawBytes: number, bodyBytes: number, cfg: BudgetConfig = DEFAULT_BUDGET,
): AdmitResult<F> {
  const accepted: F[] = [];
  const rejected: { file: F; reason: string }[] = [];
  let running = existingRawBytes;
  for (const file of files) {
    if (file.size === 0) {
      rejected.push({ file, reason: `"${file.name}" is empty (or is a folder) and was skipped.` });
      continue;
    }
    if (file.size > cfg.maxRawAttachmentBytes) {
      rejected.push({ file, reason: `"${file.name}" (${formatBytes(file.size)}) exceeds the ${formatBytes(cfg.maxRawAttachmentBytes)} per-file limit.` });
      continue;
    }
    const projected = encodedSize(running + file.size) + bodyBytes + cfg.mimeOverheadBytes;
    if (projected > cfg.maxEncodedMessageBytes) {
      rejected.push({
        file,
        reason: `Adding "${file.name}" would make the message about ${formatBytes(projected)} encoded, over Gmail's ${formatBytes(cfg.maxEncodedMessageBytes)} limit.`,
      });
      continue;
    }
    running += file.size;
    accepted.push(file);
  }
  return { accepted, rejected };
}
