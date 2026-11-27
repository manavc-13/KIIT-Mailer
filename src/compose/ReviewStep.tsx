// Step 3: Review & send. Left: summary, deliverability checklist, unsubscribe option, test send
// and the send button (with a plain confirmation dialog). Right: the full preview.
import { useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useQueryClient } from '@tanstack/react-query';
import { toast } from 'sonner';
import {
  CheckCircle2, AlertTriangle, XCircle, Info, Send, Loader2, FlaskConical, ArrowRight,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter,
  AlertDialogHeader, AlertDialogTitle,
} from '@/components/ui/alert-dialog';
import { cn } from '@/lib/utils';
import { api, errorMessage, sendErrorTitle } from '@/lib/api';
import { useAuth } from '@/lib/auth';
import { EmailPreview, renderRequestKey } from '@/preview';
import { formatBytes, isValidEmail, parseEmailList } from '../../shared/text';
import type { CreateJobRequest, Job, RenderResult, TestSendRequest } from '../../shared/types';
import { useComposeStore, type ComposeStep, type UnsubscribeOption } from './store';
import {
  autoUnsubscribeOn, firstRecipientRow, previewRows, useComposeContent, useRecipientModel, useRenderRequest, useSelectedSender,
} from './model';
import { attachmentsBlockingReason, useAttachmentMeter } from './Attachments';
import { buildChecks, quotaAfter, sortChecks, type Check, type CheckStatus } from './checks';
import { clearDraft } from './draft';
import { themeLabel, useThemes } from './ThemePicker';
import { Section, Segmented, pluralize } from './ui';

export function ReviewStep({ onGoTo }: { onGoTo: (step: ComposeStep) => void }) {
  const model = useRecipientModel();
  const mode = useComposeStore(s => s.mode);
  const unsubscribe = useComposeStore(s => s.unsubscribe);
  const attachments = useComposeStore(s => s.attachments);
  const testSent = useComposeStore(s => s.testSent);
  const previewIndex = useComposeStore(s => s.previewIndex);
  const set = useComposeStore(s => s.set);
  const content = useComposeContent();
  const { sender } = useSelectedSender();
  const meter = useAttachmentMeter();
  const themes = useThemes();
  // The last render, tagged with the request it came from. A result for any other request (sender,
  // unsubscribe or recipient changed, or the newest render failed) counts as pending, never as current.
  const [rendered, setRendered] = useState<{ key: string; result: RenderResult } | null>(null);

  const rows = useMemo(() => previewRows(model, mode), [model, mode]);
  const idx = Math.min(previewIndex, Math.max(0, rows.length - 1));
  const unsubscribeOn = unsubscribe === 'auto' ? autoUnsubscribeOn(model.build) : unsubscribe;
  const request = useRenderRequest(rows[idx], unsubscribeOn);
  const requestKey = useMemo(() => renderRequestKey(request), [request]);
  const render = rendered && rendered.key === requestKey ? rendered.result : null;

  const checks = useMemo(() => sortChecks(buildChecks({
    sender, content, mode, build: model.build, placeholders: model.placeholders,
    attachmentCount: attachments.length, attachmentBlocking: attachmentsBlockingReason(attachments),
    meter, testSent, render, renderPending: !!request && !render,
  })), [sender, content, mode, model, attachments, meter, testSent, render, request]);

  const blocking = checks.filter(c => c.status === 'fail' && c.blocking);
  const lintFails = checks.filter(c => c.status === 'fail' && !c.blocking);
  const readyAttachments = attachments.filter(a => a.status === 'ready');
  const { after } = quotaAfter(sender, model.build.quotaCost);

  return (
    <div className="grid gap-6 xl:grid-cols-[minmax(0,30rem)_minmax(0,1fr)]">
      <div className="space-y-5">
        <Section id="review-summary" title="Summary">
          <dl className="grid grid-cols-[8.5rem_minmax(0,1fr)] gap-x-4 gap-y-2.5 text-sm">
            <SummaryRow label="From">
              {sender ? <><span className="font-medium">{sender.displayName}</span> <span className="text-muted-foreground">&lt;{sender.email}&gt;</span></> : <span className="text-destructive">Not chosen</span>}
            </SummaryRow>
            <SummaryRow label="Subject"><span className="font-medium">{content.subject || <span className="text-destructive">Empty</span>}</span></SummaryRow>
            <SummaryRow label="Sending mode">{mode === 'personal' ? 'Personalized (one email each)' : 'One email, everyone in BCC'}</SummaryRow>
            <SummaryRow label="Recipients">
              {model.build.recipientCount.toLocaleString()}
              {model.build.droppedDuplicates > 0 && <span className="text-muted-foreground"> ({model.build.droppedDuplicates} duplicates removed)</span>}
            </SummaryRow>
            <SummaryRow label="Messages">
              {model.build.messageCount.toLocaleString()}
              {mode === 'bcc_broadcast' && model.build.messageCount > 0 && <span className="text-muted-foreground"> (up to 90 BCC each)</span>}
            </SummaryRow>
            <SummaryRow label="Copies">{model.build.ccTotal.toLocaleString()} CC · {model.build.bccTotal.toLocaleString()} BCC</SummaryRow>
            <SummaryRow label="Attachments">
              {readyAttachments.length === 0 ? <span className="text-muted-foreground">None</span> : (
                <ul className="space-y-0.5">
                  {readyAttachments.map(a => <li key={a.localId} className="truncate">{a.filename} <span className="text-muted-foreground">· {formatBytes(a.size)}</span></li>)}
                  <li className="text-xs text-muted-foreground">Total {formatBytes(meter.rawBytes)} (≈ {formatBytes(meter.encodedBytes)} per message)</li>
                </ul>
              )}
            </SummaryRow>
            <SummaryRow label="Design">
              {content.editorMode === 'html' ? 'Custom HTML' : `${themeLabel(content.theme, themes)} letterhead`}
              <span className="text-muted-foreground"> · signature {content.includeSignature ? 'on' : 'off'}</span>
            </SummaryRow>
            <SummaryRow label="Daily quota">
              {sender ? (
                <span className={cn(after < 0 && 'font-medium text-destructive')}>
                  {after < 0 ? `Short by ${Math.abs(after).toLocaleString()}` : `${after.toLocaleString()} remaining after this send`}
                  <span className="block text-xs text-muted-foreground">This send counts {model.build.quotaCost.toLocaleString()} (To + CC + BCC)</span>
                </span>
              ) : '—'}
            </SummaryRow>
          </dl>
        </Section>

        <Section id="review-checks" title="Deliverability checklist" description="Fix the red items before sending; the amber ones are worth a look.">
          <ul className="space-y-1.5">
            {checks.map(c => <CheckRow key={c.id} check={c} onFix={onGoTo} />)}
          </ul>
        </Section>

        <UnsubscribeCard autoOn={autoUnsubscribeOn(model.build)} />
        <TestSendCard />

        <SendCard blocking={blocking} lintFails={lintFails} />
      </div>

      <div className="min-w-0 xl:sticky xl:top-4 xl:self-start">
        <EmailPreview
          layout="full"
          request={request}
          recipientIndex={idx}
          recipientCount={rows.length}
          onRecipientIndexChange={i => set({ previewIndex: i })}
          fromName={sender?.displayName}
          fromEmail={sender?.email}
          toLabel={mode === 'bcc_broadcast' ? 'Undisclosed recipients' : rows[idx]?.Email}
          onRendered={(result, key) => setRendered({ key, result })}
        />
      </div>
    </div>
  );
}

function SummaryRow({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <>
      <dt className="text-muted-foreground">{label}</dt>
      <dd className="min-w-0 break-words">{children}</dd>
    </>
  );
}

const CHECK_ICON: Record<CheckStatus, { icon: typeof Info; cls: string; label: string }> = {
  pass: { icon: CheckCircle2, cls: 'text-success', label: 'Passed' },
  warn: { icon: AlertTriangle, cls: 'text-warning', label: 'Warning' },
  fail: { icon: XCircle, cls: 'text-destructive', label: 'Failed' },
  info: { icon: Info, cls: 'text-muted-foreground', label: 'Note' },
};

function CheckRow({ check, onFix }: { check: Check; onFix: (s: ComposeStep) => void }) {
  const { icon: Icon, cls, label } = CHECK_ICON[check.status];
  return (
    <li className={cn('flex items-start gap-2.5 rounded-md px-2.5 py-2 text-sm',
      check.status === 'fail' ? 'bg-destructive/8' : check.status === 'warn' ? 'bg-warning-soft' : '')}>
      <Icon className={cn('mt-0.5 size-4 shrink-0', cls)} aria-label={label} />
      <div className="min-w-0 flex-1">
        <div className="font-medium leading-snug">{check.title}</div>
        {check.detail && <div className="mt-0.5 text-xs text-muted-foreground">{check.detail}</div>}
      </div>
      {check.fixStep && (check.status === 'fail' || check.status === 'warn') && (
        <Button type="button" variant="ghost" size="xs" onClick={() => onFix(check.fixStep!)}>Fix <ArrowRight /></Button>
      )}
    </li>
  );
}

function UnsubscribeCard({ autoOn }: { autoOn: boolean }) {
  const unsubscribe = useComposeStore(s => s.unsubscribe);
  const set = useComposeStore(s => s.set);
  const value = unsubscribe === 'auto' ? 'auto' : unsubscribe ? 'on' : 'off';
  return (
    <Section id="review-unsubscribe" title="Unsubscribe link"
      actions={
        <Segmented<'auto' | 'on' | 'off'>
          ariaLabel="Unsubscribe header" size="sm" value={value}
          onChange={v => set({ unsubscribe: (v === 'auto' ? 'auto' : v === 'on') as UnsubscribeOption })}
          options={[{ value: 'auto', label: 'Auto' }, { value: 'on', label: 'On' }, { value: 'off', label: 'Off' }]}
        />
      }>
      <p className="text-[13px] text-muted-foreground">
        Adds a one-click unsubscribe header (shown by Gmail and Outlook next to the sender) and a footer link. Gmail expects it for bulk mail,
        and people who unsubscribe are skipped in future sends. <strong className="font-medium text-foreground">Auto</strong> turns it on when
        more than 50 recipients are outside kiit.ac.in; for this list it is <strong className="font-medium text-foreground">{autoOn ? 'on' : 'off'}</strong>.
        Leave it off for mandatory official notices to staff.
      </p>
    </Section>
  );
}

function TestSendCard() {
  const { user } = useAuth();
  const content = useComposeContent();
  const model = useRecipientModel();
  const mode = useComposeStore(s => s.mode);
  const senderAccountId = useComposeStore(s => s.senderAccountId);
  const attachments = useComposeStore(s => s.attachments);
  const set = useComposeStore(s => s.set);
  const [target, setTarget] = useState<'me' | 'others'>('me');
  const [others, setOthers] = useState('');
  const [busy, setBusy] = useState(false);
  const [results, setResults] = useState<{ to: string; success: boolean; error?: string }[] | null>(null);

  const list = target === 'me' ? (user?.email ? [user.email] : []) : parseEmailList(others);
  const invalid = list.filter(e => !isValidEmail(e));
  const tooMany = list.length > 5;
  const blockedBy = attachmentsBlockingReason(attachments);
  const disabled = busy || !senderAccountId || !list.length || invalid.length > 0 || tooMany || !!blockedBy;

  const send = async () => {
    if (!senderAccountId) return;
    setBusy(true); setResults(null);
    try {
      const body: TestSendRequest = {
        ...content,
        senderAccountId,
        to: list,
        row: firstRecipientRow(model, mode),
        attachmentIds: attachments.filter(a => a.status === 'ready' && a.ref).map(a => a.ref!.id),
      };
      const res = await api.post<{ results: { to: string; success: boolean; error?: string }[] }>('/jobs/test', body);
      setResults(res.results);
      const ok = res.results.filter(r => r.success).length;
      if (ok) { set({ testSent: true }); toast.success(`Test sent to ${pluralize(ok, 'address', 'addresses')}`); }
      if (ok < res.results.length) toast.error('Some test emails failed. See the details below.');
    } catch (e) {
      toast.error(sendErrorTitle(e, 'Test email failed'), { description: errorMessage(e) });
    } finally {
      setBusy(false);
    }
  };

  return (
    <Section id="review-test" title="Send a test" description="Uses the first recipient's details for placeholders. Check it on your phone as well.">
      <div className="space-y-3">
        <div role="radiogroup" aria-label="Test recipients" className="space-y-2">
          <label className="flex items-center gap-2 text-sm">
            <input type="radio" name="test-target" className="accent-primary" checked={target === 'me'} onChange={() => setTarget('me')} />
            To me <span className="text-muted-foreground">({user?.email})</span>
          </label>
          <label className="flex items-center gap-2 text-sm">
            <input type="radio" name="test-target" className="accent-primary" checked={target === 'others'} onChange={() => setTarget('others')} />
            To up to 5 addresses
          </label>
        </div>
        {target === 'others' && (
          <div className="space-y-1">
            <Label htmlFor="test-addresses" className="sr-only">Test addresses</Label>
            <Input id="test-addresses" autoFocus value={others} placeholder="colleague@kiit.ac.in, me@gmail.com"
              aria-invalid={invalid.length > 0 || tooMany} onChange={e => setOthers(e.target.value)} />
            {(invalid.length > 0 || tooMany) && (
              <p className="text-xs text-destructive">{tooMany ? 'At most 5 addresses.' : `Not valid: ${invalid.join(', ')}`}</p>
            )}
          </div>
        )}
        <div className="flex flex-wrap items-center gap-3">
          <Button type="button" variant="outline" onClick={send} disabled={disabled}>
            {busy ? <Loader2 className="animate-spin" /> : <FlaskConical />} Send test
          </Button>
          {blockedBy && <span className="text-xs text-muted-foreground">{blockedBy}</span>}
        </div>
        {results && (
          <ul className="space-y-1 text-sm" aria-live="polite">
            {results.map(r => (
              <li key={r.to} className="flex items-start gap-2">
                {r.success ? <CheckCircle2 className="mt-0.5 size-4 text-success" aria-label="Sent" /> : <XCircle className="mt-0.5 size-4 text-destructive" aria-label="Failed" />}
                <span><span className="font-medium">{r.to}</span>{r.error && <span className="block text-xs text-destructive">{r.error}</span>}</span>
              </li>
            ))}
          </ul>
        )}
      </div>
    </Section>
  );
}

function SendCard({ blocking, lintFails }: { blocking: Check[]; lintFails: Check[] }) {
  const navigate = useNavigate();
  const qc = useQueryClient();
  const { user } = useAuth();
  const model = useRecipientModel();
  const content = useComposeContent();
  const { sender } = useSelectedSender();
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const n = model.build.recipientCount;

  const send = async () => {
    const s = useComposeStore.getState();
    if (!s.senderAccountId) return;
    setBusy(true);
    try {
      const body: CreateJobRequest = {
        ...content,
        senderAccountId: s.senderAccountId,
        templateId: s.templateId,
        mode: s.mode,
        recipients: model.build.recipients,
        // Personal mode: cc/bcc are already resolved per recipient; the server's merge uses the same
        // rules (de-duped, never repeating To), so sending the globals too only records them on the job.
        globalCc: s.globalCc.trim() || undefined,
        globalBcc: s.globalBcc.trim() || undefined,
        attachmentIds: s.attachments.filter(a => a.status === 'ready' && a.ref).map(a => a.ref!.id),
        unsubscribe: s.unsubscribe,
        dedupe: s.mode === 'bcc_broadcast' || (s.source !== 'single' && s.dedupe),
      };
      const res = await api.post<{ job: Job; warnings: string[] }>('/jobs', body);
      (res.warnings ?? []).forEach(w => toast.warning(w, { duration: 10000 }));
      clearDraft(user?.id);
      useComposeStore.getState().reset(user?.id ?? null);
      void qc.invalidateQueries({ queryKey: ['jobs'] });
      void qc.invalidateQueries({ queryKey: ['senders'] });
      setOpen(false);
      navigate(`/send/${res.job.id}`);
    } catch (e) {
      toast.error(sendErrorTitle(e, 'Could not start sending'), { description: errorMessage(e), duration: 10000 });
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="rounded-lg border-2 border-primary/15 bg-card p-5 shadow-xs dark:border-gold/25">
      {blocking.length > 0 ? (
        <p className="mb-3 text-sm text-destructive">Fix {pluralize(blocking.length, 'item')} marked in red to send.</p>
      ) : (
        <p className="mb-3 text-sm text-muted-foreground">Sending starts on the next screen; keep that tab open until it finishes.</p>
      )}
      <Button type="button" size="lg" className="w-full" disabled={blocking.length > 0 || !n} onClick={() => setOpen(true)}>
        <Send /> Send to {pluralize(n, 'recipient')}
      </Button>

      <AlertDialog open={open} onOpenChange={o => { if (!busy) setOpen(o); }}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Send this email now?</AlertDialogTitle>
            <AlertDialogDescription asChild>
              <div className="space-y-3 text-sm">
                <dl className="grid grid-cols-[6rem_minmax(0,1fr)] gap-x-3 gap-y-1.5 rounded-md border bg-muted/50 p-3 text-foreground">
                  <dt className="text-muted-foreground">From</dt><dd className="break-words">{sender?.displayName} &lt;{sender?.email}&gt;</dd>
                  <dt className="text-muted-foreground">To</dt>
                  <dd>{pluralize(n, 'recipient')}{model.build.messageCount !== n && ` in ${pluralize(model.build.messageCount, 'message')}`}</dd>
                  <dt className="text-muted-foreground">Subject</dt><dd className="break-words">{content.subject}</dd>
                </dl>
                {lintFails.length > 0 && (
                  <p className="text-destructive">
                    The content check reported {pluralize(lintFails.length, 'problem')}: {lintFails.map(c => c.title).join('; ')}. Sending anyway may hurt delivery.
                  </p>
                )}
                <p>This cannot be undone once messages go out. You can pause or cancel from the progress screen.</p>
              </div>
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel disabled={busy}>Go back</AlertDialogCancel>
            <AlertDialogAction disabled={busy} onClick={e => { e.preventDefault(); void send(); }}>
              {busy ? <Loader2 className="animate-spin" /> : <Send />} {lintFails.length ? 'Send anyway' : `Send to ${pluralize(n, 'recipient')}`}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
