// Step 2: the message. Template quick-switch, subject / preheader, rich or HTML editor,
// theme + signature, optional plain-text part, attachments.
import { useRef, useState } from 'react';
import { Braces, ChevronDown, Code2, Save, Type, RefreshCw } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Switch } from '@/components/ui/switch';
import { Textarea } from '@/components/ui/textarea';
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuLabel, DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter,
  AlertDialogHeader, AlertDialogTitle,
} from '@/components/ui/alert-dialog';
import { cn } from '@/lib/utils';
import { EmailEditor, HtmlCodeEditor } from '@/editor';
import { TemplatePicker } from '@/templates';
import type { EditorMode, Template, ThemeId } from '../../shared/types';
import { useComposeStore } from './store';
import { useRecipientModel, useSelectedSender } from './model';
import { AttachmentsPanel } from './Attachments';
import { ThemePicker } from './ThemePicker';
import { CharCounter, Section, Segmented } from './ui';

export function ComposeStep({ onPickTemplate, onOpenSave }: { onPickTemplate: (t: Template) => void; onOpenSave: () => void }) {
  const model = useRecipientModel();
  const mode = useComposeStore(s => s.mode);
  const templateId = useComposeStore(s => s.templateId);
  const templateMeta = useComposeStore(s => s.templateMeta);
  const dirty = useComposeStore(s => s.dirty);
  const placeholders = mode === 'bcc_broadcast' ? [] : model.placeholders;

  return (
    <div className="space-y-5">
      <section aria-label="Template" className="flex flex-wrap items-center gap-3 rounded-lg border bg-card px-4 py-3 shadow-xs">
        <span className="text-sm font-medium">Template</span>
        <TemplatePicker value={templateId} onPick={onPickTemplate} className="min-w-[14rem] flex-1 sm:flex-none" />
        {templateMeta && (
          <span className="text-xs text-muted-foreground">
            {dirty ? 'Edited since loading ' : 'Loaded '}<strong className="text-foreground">{templateMeta.name}</strong>
          </span>
        )}
        <Button type="button" variant="outline" size="sm" className="ml-auto" onClick={onOpenSave} title="Ctrl+S">
          <Save /> {templateMeta?.isOwner ? 'Save / update template' : 'Save as template'}
        </Button>
      </section>

      <SubjectFields placeholders={placeholders} />
      <BodyEditor placeholders={placeholders} />
      <PlainTextSection />
      <AttachmentsPanel />
    </div>
  );
}

// ------------------------------------------------------------------ subject + preheader
function insertAt(el: HTMLInputElement | null, current: string, token: string): { value: string; caret: number } {
  const start = el?.selectionStart ?? current.length;
  const end = el?.selectionEnd ?? current.length;
  return { value: current.slice(0, start) + token + current.slice(end), caret: start + token.length };
}

function PlaceholderMenu({ placeholders, onInsert, label }: { placeholders: string[]; onInsert: (p: string) => void; label: string }) {
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button type="button" variant="ghost" size="xs" aria-label={label} disabled={!placeholders.length}
          title={placeholders.length ? label : 'No placeholders: add recipients with columns first'}>
          <Braces /> Insert field <ChevronDown />
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="max-h-72 overflow-y-auto">
        <DropdownMenuLabel className="text-xs font-normal text-muted-foreground">Replaced per recipient</DropdownMenuLabel>
        {placeholders.map(p => <DropdownMenuItem key={p} onSelect={() => onInsert(p)}><code>{`{${p}}`}</code></DropdownMenuItem>)}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

function SubjectFields({ placeholders }: { placeholders: string[] }) {
  const subject = useComposeStore(s => s.subject);
  const preheader = useComposeStore(s => s.preheader);
  const edit = useComposeStore(s => s.edit);
  const subjectRef = useRef<HTMLInputElement>(null);
  const preRef = useRef<HTMLInputElement>(null);

  const insert = (which: 'subject' | 'preheader', p: string) => {
    const el = which === 'subject' ? subjectRef.current : preRef.current;
    const cur = which === 'subject' ? subject : preheader;
    const { value, caret } = insertAt(el, cur, `{${p}}`);
    edit({ [which]: value });
    requestAnimationFrame(() => { el?.focus(); el?.setSelectionRange(caret, caret); });
  };

  return (
    <Section id="compose-subject" title="Subject and inbox preview">
      <div className="space-y-4">
        <div className="space-y-1.5">
          <div className="flex items-center justify-between gap-2">
            <Label htmlFor="subject">Subject <span className="text-destructive" aria-hidden>*</span></Label>
            <div className="flex items-center gap-2">
              <PlaceholderMenu placeholders={placeholders} label="Insert a field into the subject" onInsert={p => insert('subject', p)} />
              <CharCounter count={subject.length} warnAbove={60} dangerAbove={78} />
            </div>
          </div>
          <Input id="subject" ref={subjectRef} value={subject} maxLength={300} placeholder="e.g. Invitation: IQAC review meeting, 12 December"
            onChange={e => edit({ subject: e.target.value })} aria-describedby="subject-hint" />
          <p id="subject-hint" className="text-xs text-muted-foreground">Keep it under about 60 characters so phones show it in full. Avoid ALL CAPS and "!!!".</p>
        </div>
        <div className="space-y-1.5">
          <div className="flex items-center justify-between gap-2">
            <Label htmlFor="preheader">Preheader <span className="font-normal text-muted-foreground">(optional)</span></Label>
            <div className="flex items-center gap-2">
              <PlaceholderMenu placeholders={placeholders} label="Insert a field into the preheader" onInsert={p => insert('preheader', p)} />
              <CharCounter count={preheader.length} warnAbove={110} warnBelow={30} />
            </div>
          </div>
          <Input id="preheader" ref={preRef} value={preheader} maxLength={250} placeholder="One line that appears after the subject in the inbox list"
            onChange={e => edit({ preheader: e.target.value })} aria-describedby="preheader-hint" />
          <p id="preheader-hint" className="text-xs text-muted-foreground">
            Shown next to the subject in most inboxes. 40–100 characters works best; if left empty, clients show the first words of the body.
          </p>
        </div>
      </div>
    </Section>
  );
}

// ------------------------------------------------------------------ body editor
function extractBody(html: string): string {
  const m = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
  return (m ? m[1] : html).replace(/<!--[\s\S]*?-->/g, '').trim();
}

function wrapAsDocument(fragment: string, subject: string): string {
  const title = subject.replace(/[<>&]/g, '');
  return `<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>${title}</title>\n</head>\n<body>\n${fragment}\n</body>\n</html>\n`;
}

function BodyEditor({ placeholders }: { placeholders: string[] }) {
  const editorMode = useComposeStore(s => s.editorMode);
  const editorKey = useComposeStore(s => s.editorKey);
  const bodyHtml = useComposeStore(s => s.bodyHtml);
  const theme = useComposeStore(s => s.theme);
  const includeSignature = useComposeStore(s => s.includeSignature);
  const edit = useComposeStore(s => s.edit);
  const set = useComposeStore(s => s.set);
  const { sender } = useSelectedSender();
  const [pendingMode, setPendingMode] = useState<EditorMode | null>(null);

  const hasBody = bodyHtml.replace(/<[^>]+>/g, '').trim().length > 0;

  const switchMode = (to: EditorMode) => {
    if (to === editorMode) return;
    if (hasBody) setPendingMode(to);
    else applyMode(to);
  };

  const applyMode = (to: EditorMode) => {
    const s = useComposeStore.getState();
    if (to === 'html') {
      const seeded = s.bodyHtml.trim() ? wrapAsDocument(s.bodyHtml, s.subject) : '';
      set({ editorMode: 'html', doc: null, seedHtml: null, bodyHtml: seeded, editorKey: s.editorKey + 1, dirty: true });
    } else {
      const fragment = extractBody(s.bodyHtml);
      set({ editorMode: 'rich', doc: null, seedHtml: fragment, bodyHtml: fragment, editorKey: s.editorKey + 1, dirty: true });
    }
  };

  return (
    <Section
      id="compose-body"
      title="Message"
      description={editorMode === 'rich'
        ? 'Write as you would in a document; the theme styles it like the final email.'
        : 'A complete HTML email document. It is sent as written (no letterhead); the signature is added if switched on.'}
      actions={
        <Segmented<EditorMode>
          ariaLabel="Editor mode" size="sm" value={editorMode} onChange={switchMode}
          options={[{ value: 'rich', label: 'Rich text', icon: Type }, { value: 'html', label: 'HTML', icon: Code2 }]}
        />
      }
    >
      {editorMode === 'rich' && <ThemePicker />}

      <div className="mt-1">
        {editorMode === 'rich' ? (
          <RichEditorHost key={`rich-${editorKey}`} theme={theme} placeholders={placeholders} />
        ) : (
          <HtmlCodeEditor
            key={`html-${editorKey}`}
            value={bodyHtml}
            placeholders={placeholders}
            minHeight={420}
            onChange={html => edit({ bodyHtml: html })}
          />
        )}
      </div>

      <div className="mt-4 flex flex-wrap items-center justify-between gap-3 border-t pt-4">
        <div className="flex items-center gap-3">
          <Switch id="include-signature" checked={includeSignature} onCheckedChange={v => edit({ includeSignature: v })} />
          <Label htmlFor="include-signature" className="font-normal">
            Include signature{sender ? <> of <strong className="font-medium">{sender.displayName}</strong></> : ''}
          </Label>
        </div>
        {sender && !sender.signatureHtml && includeSignature && (
          <span className="text-xs text-muted-foreground">This account has no signature set yet; ask the administrator to add one.</span>
        )}
      </div>

      <AlertDialog open={pendingMode !== null} onOpenChange={o => { if (!o) setPendingMode(null); }}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Switch to {pendingMode === 'html' ? 'HTML' : 'rich text'}?</AlertDialogTitle>
            <AlertDialogDescription>
              {pendingMode === 'html'
                ? 'Your message will be converted to HTML source you can edit directly. The letterhead theme no longer applies in HTML mode, and switching back later may lose layout.'
                : 'The HTML will be imported into the rich-text editor. Complex layouts (tables, custom styles, scripts) can be simplified or lost.'}
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Stay in {editorMode === 'html' ? 'HTML' : 'rich text'}</AlertDialogCancel>
            <AlertDialogAction onClick={() => { if (pendingMode) applyMode(pendingMode); setPendingMode(null); }}>
              <RefreshCw /> Convert and switch
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </Section>
  );
}

/** Mounted per editorKey: hands the editor its initial content once, then only listens. */
function RichEditorHost({ theme, placeholders }: { theme: ThemeId; placeholders: string[] }) {
  const [initial] = useState(() => {
    const s = useComposeStore.getState();
    return { doc: s.doc, html: s.seedHtml ?? (s.doc ? undefined : s.bodyHtml || undefined) };
  });
  return (
    <EmailEditor
      doc={initial.doc}
      initialHtml={initial.html}
      theme={theme}
      placeholders={placeholders}
      minHeight={360}
      onChange={v => {
        const cur = useComposeStore.getState();
        // Editors often emit once on mount; an unchanged body must not mark the compose dirty.
        if (v.html === cur.bodyHtml) cur.set({ doc: v.doc });
        else cur.edit({ doc: v.doc, bodyHtml: v.html, seedHtml: null });
      }}
    />
  );
}

// ------------------------------------------------------------------ plain text
function PlainTextSection() {
  const enabled = useComposeStore(s => s.customTextEnabled);
  const text = useComposeStore(s => s.customText);
  const edit = useComposeStore(s => s.edit);
  const [open, setOpen] = useState(enabled);
  return (
    <section className="rounded-lg border bg-card shadow-xs">
      <button
        type="button" aria-expanded={open} aria-controls="plain-text-body" onClick={() => setOpen(o => !o)}
        className="flex w-full items-center justify-between gap-3 rounded-lg px-5 py-3.5 text-left outline-none focus-visible:ring-2 focus-visible:ring-ring"
      >
        <span>
          <span className="block font-serif text-[17px] font-semibold">Plain-text version <span className="font-sans text-sm font-normal text-muted-foreground">(optional)</span></span>
          <span className="block text-[13px] text-muted-foreground">
            {enabled && text.trim() ? 'Using your own plain-text part.' : 'Generated automatically from the message when left empty.'}
          </span>
        </span>
        <ChevronDown className={cn('size-4 shrink-0 transition-transform', open && 'rotate-180')} />
      </button>
      {open && (
        <div id="plain-text-body" className="space-y-3 border-t px-5 py-4">
          <div className="flex items-center gap-3">
            <Switch id="custom-text" checked={enabled} onCheckedChange={v => edit({ customTextEnabled: v })} />
            <Label htmlFor="custom-text" className="font-normal">Write my own plain-text part</Label>
          </div>
          {enabled && (
            <Textarea aria-label="Plain-text message" rows={8} value={text} className="font-mono text-[13px]"
              placeholder={'Dear {Name},\n\n…'} onChange={e => edit({ customText: e.target.value })} />
          )}
          <p className="text-xs text-muted-foreground">
            Some mail apps and screen readers show only the plain-text part. Placeholders work here too.
          </p>
        </div>
      )}
    </section>
  );
}
