// Compose: the 3-step flow Recipients -> Compose -> Review & send, with the From selector in the
// header and a docked live preview on large screens.
import { useCallback, useEffect, useMemo, useState } from 'react';
import { useSearchParams } from 'react-router-dom';
import { toast } from 'sonner';
import { ArrowLeft, ArrowRight, Check, Eye, EyeOff, KeyRound, Loader2 } from 'lucide-react';
import { PageBody, PageHeader } from '@/components/AppShell';
import { Button } from '@/components/ui/button';
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter,
  AlertDialogHeader, AlertDialogTitle,
} from '@/components/ui/alert-dialog';
import { cn } from '@/lib/utils';
import { api, errorMessage } from '@/lib/api';
import { useAuth } from '@/lib/auth';
import { EmailPreview } from '@/preview';
import { SaveTemplateDialog } from '@/templates';
import type { Template } from '../../shared/types';
import { STEPS, hasComposeContent, useComposeStore, type ComposeStep } from '@/compose/store';
import { clearDraft, useDraftAutosave } from '@/compose/draft';
import {
  previewRows, useRecipientModel, useRenderRequest, useSaveTemplateContent, useSelectedSender, type RecipientModel,
} from '@/compose/model';
import { SenderSelect } from '@/compose/SenderSelect';
import { RecipientsStep } from '@/compose/RecipientsStep';
import { ComposeStep as ComposeStepView } from '@/compose/ComposeStep';
import { ReviewStep } from '@/compose/ReviewStep';

function useMediaQuery(query: string): boolean {
  const [matches, setMatches] = useState(() => typeof window !== 'undefined' && window.matchMedia(query).matches);
  useEffect(() => {
    const mq = window.matchMedia(query);
    const on = () => setMatches(mq.matches);
    on();
    mq.addEventListener('change', on);
    return () => mq.removeEventListener('change', on);
  }, [query]);
  return matches;
}

const STEP_INDEX: Record<ComposeStep, number> = { recipients: 0, compose: 1, review: 2 };

/** Why the user can't leave `step` going forward, or null. */
function stepBlocker(step: ComposeStep, model: RecipientModel): string | null {
  const s = useComposeStore.getState();
  if (step === 'recipients') {
    if (s.source !== 'single' && model.rows.length && !model.roleMap.email) return 'Map the Email column first.';
    if (!model.build.recipientCount) {
      return s.source === 'single' ? 'Enter a valid recipient email address.' : 'Add at least one recipient with a valid email address.';
    }
  }
  if (step === 'compose') {
    if (!s.subject.trim()) return 'Add a subject.';
    const text = s.bodyHtml.replace(/<style[\s\S]*?<\/style>/gi, '').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
    if (!text && !/<img\b/i.test(s.bodyHtml)) return 'Write the message first.';
  }
  return null;
}

export default function ComposePage() {
  const { user } = useAuth();
  useDraftAutosave(user?.id);

  const step = useComposeStore(s => s.step);
  const setStep = useComposeStore(s => s.setStep);
  const { sender, senders, loading: sendersLoading, error: sendersError } = useSelectedSender();
  const model = useRecipientModel();
  const saveContent = useSaveTemplateContent();
  const templateMeta = useComposeStore(s => s.templateMeta);
  const [saveOpen, setSaveOpen] = useState(false);
  const [pendingTemplate, setPendingTemplate] = useState<Template | null>(null);
  const [params, setParams] = useSearchParams();
  const [loadingTemplate, setLoadingTemplate] = useState(false);
  const isLarge = useMediaQuery('(min-width: 1024px)');

  // ---- sender auto-selection (after the per-user store reset / draft restore has happened)
  const ownerId = useComposeStore(s => s.ownerId);
  const senderAccountId = useComposeStore(s => s.senderAccountId);
  useEffect(() => {
    if (sendersLoading || !user || ownerId !== user.id) return;
    const s = useComposeStore.getState();
    if (senders.some(x => x.id === senderAccountId)) return;
    if (senders.length === 1) s.selectSender(senders[0]);
    else if (senderAccountId) s.set({ senderAccountId: null });
  }, [senders, sendersLoading, user, ownerId, senderAccountId]);

  // ---- templates
  const applyTemplate = useCallback((t: Template) => {
    const ids = senders.map(s => s.id);
    useComposeStore.getState().applyTemplate(t, ids);
    const pinned = t.defaultSenderAccountId && ids.includes(t.defaultSenderAccountId)
      ? senders.find(s => s.id === t.defaultSenderAccountId) : null;
    toast.success(`Loaded "${t.name}"`, { description: pinned ? `From set to ${pinned.displayName}` : undefined });
  }, [senders]);

  const requestTemplate = useCallback((t: Template) => {
    const s = useComposeStore.getState();
    if (s.templateId === t.id && !s.dirty) return;
    if (s.dirty && hasComposeContent(s)) setPendingTemplate(t);
    else applyTemplate(t);
  }, [applyTemplate]);

  // ?template=<id> (e.g. "Use" on the Templates page)
  const templateParam = params.get('template');
  useEffect(() => {
    if (!templateParam || !user || sendersLoading || ownerId !== user.id) return;
    let cancelled = false;
    setLoadingTemplate(true);
    api.get<{ template: Template }>(`/templates/${encodeURIComponent(templateParam)}`)
      .then(({ template }) => {
        if (cancelled) return;
        useComposeStore.getState().setStep('compose');
        requestTemplate(template);
      })
      .catch(e => { if (!cancelled) toast.error('Could not open that template', { description: errorMessage(e) }); })
      .finally(() => {
        if (cancelled) return;
        setLoadingTemplate(false);
        setParams(p => { p.delete('template'); return p; }, { replace: true });
      });
    return () => { cancelled = true; };
  }, [templateParam, user, ownerId, sendersLoading, requestTemplate, setParams]);

  // ?new=1 ("New template" on the Templates page): start from a blank compose. If there is
  // unsaved work, don't silently discard it; offer the reset as an explicit toast action.
  const newParam = params.get('new');
  useEffect(() => {
    if (!newParam || !user || ownerId !== user.id) return;
    setParams(p => { p.delete('new'); return p; }, { replace: true });
    const startNew = () => {
      clearDraft(user.id);
      useComposeStore.getState().reset(user.id);
      useComposeStore.getState().setStep('compose');
      toast.success('Started a new email', { description: 'Write it, then use "Save as template" to add it to your library.' });
    };
    const s = useComposeStore.getState();
    if (hasComposeContent(s)) {
      toast('You have an email in progress', {
        description: 'Starting a new template will discard it.',
        action: { label: 'Discard and start new', onClick: startNew },
        duration: 10000,
      });
    } else startNew();
  }, [newParam, user, ownerId, setParams]);

  // ---- step navigation
  const goTo = useCallback((target: ComposeStep) => {
    const from = STEP_INDEX[useComposeStore.getState().step];
    const to = STEP_INDEX[target];
    if (to > from) {
      for (let i = from; i < to; i++) {
        const blocker = stepBlocker(STEPS[i].id, model);
        if (blocker) { toast.warning(blocker); if (i !== from) setStep(STEPS[i].id); return; }
      }
    }
    setStep(target);
    document.querySelector('main')?.scrollTo({ top: 0, behavior: 'smooth' });
  }, [model, setStep]);

  const next = useCallback(() => {
    const i = STEP_INDEX[useComposeStore.getState().step];
    if (i < STEPS.length - 1) goTo(STEPS[i + 1].id);
  }, [goTo]);

  // ---- keyboard: Ctrl+Enter next step, Ctrl+S save template
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      if (e.key === 'Enter') { e.preventDefault(); next(); }
      else if (e.key.toLowerCase() === 's' && !e.shiftKey && !e.altKey) { e.preventDefault(); setSaveOpen(true); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [next]);

  // ---- no sender accounts at all
  if (!sendersLoading && !sendersError && senders.length === 0) {
    return (
      <>
        <PageHeader title="Compose" />
        <PageBody>
          <div className="mx-auto mt-10 max-w-lg rounded-lg border bg-card p-8 text-center shadow-xs">
            <KeyRound className="mx-auto size-8 text-gold" strokeWidth={1.5} />
            <h2 className="mt-4 text-xl font-semibold">No sender account yet</h2>
            <p className="mt-2 text-sm text-muted-foreground">
              Ask the administrator to give you access to a sender account. Once an @kiit.ac.in account is allotted to you, it appears here.
            </p>
          </div>
        </PageBody>
      </>
    );
  }

  return (
    <>
      <PageHeader
        title="Compose"
        description={sender
          ? <>Signature and letterhead from <strong className="font-medium text-foreground">{sender.displayName}</strong>.</>
          : 'Choose the account to send from.'}
        actions={<SenderSelect />}
      />
      {sendersError ? (
        <div role="alert" className="mx-6 mt-4 rounded-md border border-destructive/30 bg-destructive/5 px-4 py-2 text-sm text-destructive lg:mx-10">
          Could not load your sender accounts: {errorMessage(sendersError)}
        </div>
      ) : null}
      <StepBar current={step} onGoTo={goTo} />
      <PageBody className="pt-5">
        {loadingTemplate && (
          <div className="mb-4 flex items-center gap-2 text-sm text-muted-foreground"><Loader2 className="size-4 animate-spin" /> Opening template…</div>
        )}
        {step === 'review' ? (
          <ReviewStep onGoTo={goTo} />
        ) : (
          <div className="grid gap-6 lg:grid-cols-[minmax(0,1fr)_minmax(340px,420px)]">
            <div className="min-w-0">
              {step === 'recipients' ? <RecipientsStep /> : <ComposeStepView onPickTemplate={requestTemplate} onOpenSave={() => setSaveOpen(true)} />}
              <StepNav step={step} onBack={() => goTo(STEPS[STEP_INDEX[step] - 1].id)} onNext={next} />
              {!isLarge && <InlinePreview />}
            </div>
            {isLarge && (
              <aside aria-label="Live preview" className="min-w-0">
                <div className="sticky top-4">
                  <PreviewPanel layout="rail" />
                </div>
              </aside>
            )}
          </div>
        )}
        {step === 'review' && (
          <div className="mt-6">
            <Button type="button" variant="ghost" onClick={() => goTo('compose')}><ArrowLeft /> Back to compose</Button>
          </div>
        )}
      </PageBody>

      <SaveTemplateDialog
        open={saveOpen}
        onOpenChange={setSaveOpen}
        content={saveContent}
        existing={templateMeta}
        // SaveTemplateDialog shows the success toast ("Saved "X" to My templates" / "Updated "X"").
        onSaved={t => useComposeStore.getState().setTemplateMeta(t)}
      />

      <AlertDialog open={pendingTemplate !== null} onOpenChange={o => { if (!o) setPendingTemplate(null); }}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Replace your current message?</AlertDialogTitle>
            <AlertDialogDescription>
              Loading "{pendingTemplate?.name}" replaces the subject, preheader, body, theme and plain-text part. Recipients and attachments are kept.
              Save your current message as a template first if you want to keep it.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Keep my message</AlertDialogCancel>
            <AlertDialogAction onClick={() => { if (pendingTemplate) applyTemplate(pendingTemplate); setPendingTemplate(null); }}>
              Load template
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </>
  );
}

// ------------------------------------------------------------------ step bar + nav
function StepBar({ current, onGoTo }: { current: ComposeStep; onGoTo: (s: ComposeStep) => void }) {
  const ci = STEP_INDEX[current];
  return (
    <nav aria-label="Compose steps" className="border-b bg-card px-6 lg:px-10">
      <ol className="flex items-stretch gap-1 overflow-x-auto">
        {STEPS.map((s, i) => {
          const done = i < ci;
          const active = i === ci;
          return (
            <li key={s.id} className="flex items-center">
              <button
                type="button"
                onClick={() => onGoTo(s.id)}
                aria-current={active ? 'step' : undefined}
                className={cn(
                  'flex items-center gap-2.5 border-b-2 px-3 py-3 text-sm font-medium whitespace-nowrap transition-colors outline-none focus-visible:bg-muted',
                  active ? 'border-gold text-foreground' : 'border-transparent text-muted-foreground hover:text-foreground',
                )}
              >
                <span className={cn(
                  'grid size-6 place-items-center rounded-full text-xs font-semibold',
                  active ? 'bg-primary text-primary-foreground' : done ? 'bg-success-soft text-success ring-1 ring-success/30' : 'bg-muted text-muted-foreground ring-1 ring-border',
                )}>
                  {done ? <Check className="size-3.5" /> : i + 1}
                </span>
                {/* On phones only the current step shows its label, so all three steps fit. */}
                <span className={cn(!active && 'sr-only sm:not-sr-only')}>{s.label}</span>
              </button>
              {i < STEPS.length - 1 && <span aria-hidden className="mx-1 h-px w-6 bg-border sm:w-10" />}
            </li>
          );
        })}
      </ol>
    </nav>
  );
}

function StepNav({ step, onBack, onNext }: { step: ComposeStep; onBack: () => void; onNext: () => void }) {
  const i = STEP_INDEX[step];
  return (
    <div className="mt-6 flex items-center justify-between gap-3 border-t pt-5">
      {i > 0 ? <Button type="button" variant="ghost" onClick={onBack}><ArrowLeft /> {STEPS[i - 1].label}</Button> : <span />}
      <div className="flex items-center gap-3">
        <kbd className="hidden rounded border bg-muted px-1.5 py-0.5 font-mono text-[11px] text-muted-foreground sm:inline">Ctrl + Enter</kbd>
        <Button type="button" onClick={onNext}>Next: {STEPS[i + 1].label} <ArrowRight /></Button>
      </div>
    </div>
  );
}

// ------------------------------------------------------------------ preview
function PreviewPanel({ layout }: { layout: 'rail' | 'full' }) {
  const model = useRecipientModel();
  const mode = useComposeStore(s => s.mode);
  const previewIndex = useComposeStore(s => s.previewIndex);
  const set = useComposeStore(s => s.set);
  const { sender } = useSelectedSender();
  const rows = useMemo(() => previewRows(model, mode), [model, mode]);
  const idx = Math.min(previewIndex, Math.max(0, rows.length - 1));
  const request = useRenderRequest(rows[idx]);

  if (!sender) {
    return (
      <div className="rounded-lg border bg-card p-6 text-sm text-muted-foreground shadow-xs">
        Choose a From account to see the live preview with its letterhead and signature.
      </div>
    );
  }
  return (
    <EmailPreview
      layout={layout}
      request={request}
      recipientIndex={idx}
      recipientCount={rows.length}
      onRecipientIndexChange={i => set({ previewIndex: i })}
      fromName={sender.displayName}
      fromEmail={sender.email}
      toLabel={mode === 'bcc_broadcast' ? 'Undisclosed recipients' : rows[idx]?.Email}
    />
  );
}

/** Below lg there is no docked rail; offer the preview inline on demand. */
function InlinePreview() {
  const [open, setOpen] = useState(false);
  return (
    <div className="mt-6">
      <Button type="button" variant="outline" className="w-full" aria-expanded={open} onClick={() => setOpen(o => !o)}>
        {open ? <EyeOff /> : <Eye />} {open ? 'Hide preview' : 'Show preview'}
      </Button>
      {open && <div className="mt-4"><PreviewPanel layout="rail" /></div>}
    </div>
  );
}
