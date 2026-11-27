// Attachments: drop zone + list with per-file upload progress, removal and the encoded-size meter.
// Files upload once (uploadFile(file, 'attachment')) and every message references them by id.
import { useEffect, useMemo, useRef, useState } from 'react';
import { toast } from 'sonner';
import { Paperclip, X, Loader2, AlertTriangle, FileText, RotateCw } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { cn } from '@/lib/utils';
import { useUploadConfig } from '@/lib/queries';
import { uploadFile, deleteUpload } from '@/lib/uploads';
import { errorMessage } from '@/lib/api';
import { formatBytes } from '../../shared/text';
import { useComposeStore, newId, type AttachmentItem } from './store';
import { admitFiles, computeMeter, utf8Bytes, DEFAULT_BUDGET, type BudgetConfig } from './attachmentMath';
import { Meter, Section } from './ui';

// File objects can't live in the (persisted) store; keep them here so a failed upload can be retried.
const pendingFiles = new Map<string, File>();
const controllers = new Map<string, AbortController>();

export function useBudget(): BudgetConfig {
  const { data } = useUploadConfig();
  return data ?? DEFAULT_BUDGET;
}

export function useAttachmentMeter() {
  const attachments = useComposeStore(s => s.attachments);
  const bodyHtml = useComposeStore(s => s.bodyHtml);
  const customText = useComposeStore(s => s.customText);
  const cfg = useBudget();
  return useMemo(() => {
    const raw = attachments.filter(a => a.status !== 'error').reduce((n, a) => n + a.size, 0);
    return computeMeter(raw, utf8Bytes(bodyHtml, customText), cfg);
  }, [attachments, bodyHtml, customText, cfg]);
}

/** Reason the attachments block sending, or null. */
export function attachmentsBlockingReason(items: AttachmentItem[]): string | null {
  const uploading = items.filter(a => a.status === 'uploading').length;
  if (uploading) return `${uploading === 1 ? '1 attachment is' : `${uploading} attachments are`} still uploading.`;
  const failed = items.filter(a => a.status === 'error');
  if (failed.length) return `Remove or retry the failed attachment${failed.length > 1 ? 's' : ''}: ${failed.map(a => a.filename).join(', ')}.`;
  return null;
}

function startUpload(localId: string, file: File) {
  const { updateAttachment } = useComposeStore.getState();
  const ctrl = new AbortController();
  controllers.set(localId, ctrl);
  updateAttachment(localId, { status: 'uploading', progress: 0, error: null });
  uploadFile(file, 'attachment', {
    filename: file.name,
    signal: ctrl.signal,
    onProgress: p => updateAttachment(localId, { progress: p.total ? p.loaded / p.total : 0 }),
  }).then(ref => {
    controllers.delete(localId);
    const still = useComposeStore.getState().attachments.some(a => a.localId === localId);
    if (!still) { void deleteUpload(ref.id).catch(() => {}); return; } // removed while uploading
    pendingFiles.delete(localId);
    updateAttachment(localId, { status: 'ready', progress: 1, ref, size: ref.size || file.size });
  }).catch(err => {
    controllers.delete(localId);
    if (ctrl.signal.aborted) return;
    const msg = errorMessage(err);
    updateAttachment(localId, { status: 'error', error: msg });
    toast.error(`Could not attach "${file.name}"`, { description: msg });
  });
}

export function AttachmentsPanel() {
  const attachments = useComposeStore(s => s.attachments);
  const step = useComposeStore(s => s.step);
  const addAttachment = useComposeStore(s => s.addAttachment);
  const removeAttachment = useComposeStore(s => s.removeAttachment);
  const bodyHtml = useComposeStore(s => s.bodyHtml);
  const customText = useComposeStore(s => s.customText);
  const cfg = useBudget();
  const { data: uploadCfg } = useUploadConfig();
  const meter = useAttachmentMeter();
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const addFiles = (files: File[]) => {
    if (!files.length) return;
    const existing = useComposeStore.getState().attachments.filter(a => a.status !== 'error').reduce((n, a) => n + a.size, 0);
    const { accepted, rejected } = admitFiles(files, existing, utf8Bytes(bodyHtml, customText), cfg);
    rejected.forEach(r => toast.error(r.reason));
    for (const file of accepted) {
      const localId = newId('att');
      pendingFiles.set(localId, file);
      addAttachment({
        localId, filename: file.name, size: file.size, contentType: file.type || 'application/octet-stream',
        status: 'uploading', progress: 0, ref: null, error: null,
      });
      startUpload(localId, file);
    }
  };

  const remove = (a: AttachmentItem) => {
    controllers.get(a.localId)?.abort();
    controllers.delete(a.localId);
    pendingFiles.delete(a.localId);
    removeAttachment(a.localId);
    if (a.ref) void deleteUpload(a.ref.id).catch(() => { /* best effort */ });
  };

  const retry = (a: AttachmentItem) => {
    const f = pendingFiles.get(a.localId);
    if (f) startUpload(a.localId, f);
    else { removeAttachment(a.localId); toast('Pick the file again to retry.'); }
  };

  // Page-wide file drop while on the Compose step, so a file released slightly outside the
  // zone doesn't make the browser navigate away to it (and wipe the session).
  const addRef = useRef(addFiles);
  addRef.current = addFiles;
  useEffect(() => {
    if (step !== 'compose') return;
    let depth = 0;
    const isFiles = (e: DragEvent) => !!e.dataTransfer && Array.from(e.dataTransfer.types).includes('Files');
    const enter = (e: DragEvent) => { if (isFiles(e)) { depth++; setDragging(true); } };
    const leave = (e: DragEvent) => { if (isFiles(e)) { depth = Math.max(0, depth - 1); if (!depth) setDragging(false); } };
    const over = (e: DragEvent) => { if (isFiles(e)) { e.preventDefault(); if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy'; } };
    const drop = (e: DragEvent) => {
      depth = 0; setDragging(false);
      if (!isFiles(e) || e.defaultPrevented) return;
      e.preventDefault();
      addRef.current(Array.from(e.dataTransfer?.files ?? []));
    };
    window.addEventListener('dragenter', enter);
    window.addEventListener('dragleave', leave);
    window.addEventListener('dragover', over);
    window.addEventListener('drop', drop);
    return () => {
      window.removeEventListener('dragenter', enter);
      window.removeEventListener('dragleave', leave);
      window.removeEventListener('dragover', over);
      window.removeEventListener('drop', drop);
    };
  }, [step]);

  return (
    <Section id="compose-attachments" title="Attachments"
      description={`Uploaded once and attached to every message. Gmail's limit is ${formatBytes(cfg.maxEncodedMessageBytes)} per message after encoding (about ${formatBytes(cfg.maxRawAttachmentBytes)} of files).`}>
      {uploadCfg?.degradedReason && (
        <p className="mb-3 rounded-md border border-warning/30 bg-warning-soft px-3 py-2 text-xs text-warning">{uploadCfg.degradedReason}</p>
      )}
      <div
        role="button" tabIndex={0}
        aria-label="Attach files: drop them here or press Enter to browse"
        onClick={() => inputRef.current?.click()}
        onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inputRef.current?.click(); } }}
        className={cn(
          'flex cursor-pointer items-center justify-center gap-2 rounded-md border-2 border-dashed px-4 py-5 text-sm transition-colors outline-none focus-visible:ring-2 focus-visible:ring-ring',
          dragging ? 'border-gold bg-gold-soft text-foreground' : 'border-border bg-muted/40 text-muted-foreground hover:bg-muted',
        )}
      >
        <Paperclip className="size-4" />
        {dragging ? 'Drop files anywhere to attach' : 'Drop files here, or click to choose'}
        <input ref={inputRef} type="file" multiple className="hidden"
          onChange={e => { addFiles(Array.from(e.target.files ?? [])); e.target.value = ''; }} />
      </div>

      {attachments.length > 0 && (
        <>
          <ul className="mt-3 divide-y rounded-md border" aria-label="Attached files">
            {attachments.map(a => (
              <li key={a.localId} className={cn('flex items-center gap-3 px-3 py-2', a.status === 'error' && 'bg-destructive/5')}>
                {a.status === 'uploading' ? <Loader2 className="size-4 shrink-0 animate-spin text-muted-foreground" />
                  : a.status === 'error' ? <AlertTriangle className="size-4 shrink-0 text-destructive" />
                    : <FileText className="size-4 shrink-0 text-muted-foreground" />}
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm font-medium">{a.filename}</div>
                  <div className="text-xs text-muted-foreground">
                    {formatBytes(a.size)}
                    {a.status === 'uploading' && ` · uploading ${Math.round(a.progress * 100)}%`}
                    {a.status === 'error' && <span className="text-destructive"> · {a.error || 'upload failed'}</span>}
                  </div>
                  {a.status === 'uploading' && <Meter className="mt-1" percent={a.progress * 100} tone="primary" label={`Uploading ${a.filename}`} />}
                </div>
                {a.status === 'error' && (
                  <Button type="button" variant="ghost" size="icon-sm" aria-label={`Retry ${a.filename}`} onClick={() => retry(a)}><RotateCw /></Button>
                )}
                <Button type="button" variant="ghost" size="icon-sm" aria-label={`Remove ${a.filename}`} onClick={() => remove(a)}><X /></Button>
              </li>
            ))}
          </ul>
          <div className="mt-3 space-y-1.5">
            <Meter percent={meter.percent} tone={meter.level} label="Message size against Gmail's limit" />
            <div className={cn('text-xs', meter.level === 'danger' ? 'text-destructive' : 'text-muted-foreground')}>{meter.label}</div>
            {meter.largeWarning && (
              <p className="text-xs text-warning">
                Large attachments: some recipients' mail servers reject messages this size even though Gmail accepts it. Consider sharing a Drive link instead.
              </p>
            )}
          </div>
        </>
      )}
    </Section>
  );
}
