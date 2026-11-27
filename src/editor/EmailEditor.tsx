// <EmailEditor>: TipTap rich-text editor for email bodies and sender signatures.
// The canvas is styled by the selected theme's body CSS so writing looks like the sent mail.
import { useCallback, useEffect, useId, useMemo, useRef, useState } from 'react';
import { EditorContent, useEditor, type Editor } from '@tiptap/react';
import { Extension, type JSONContent } from '@tiptap/core';
import { EditorState } from '@tiptap/pm/state';
import { DOMParser as PMDOMParser, Slice } from '@tiptap/pm/model';
import { toast } from 'sonner';
import { FONT_CATALOG } from '../../shared/fonts';
import { themeCss, themeWebfontUrl, EMAIL_WIDTH, BODY_PADDING_X } from '../../shared/themes';
import type { ThemeId } from '../../shared/types';
import { uploadFile, rehostImage } from '@/lib/uploads';
import { cn } from '@/lib/utils';
import { buildExtensions } from './extensions';
import { insertImageFiles } from './extensions/image-uploads';
import { normalizePastedHtml, plainTextToHtml, prepareHtmlForEditor } from './paste';
import { serializeEmail } from './serialize';
import { EditorToolbar } from './toolbar';
import { ButtonDialog, ImageDialog, LinkDialog, type ButtonAttrs } from './dialogs';
import './editor.css';

export interface EmailEditorValue {
  /** TipTap/ProseMirror JSON: the source of truth that gets saved. */
  doc: unknown;
  /** Email-safe body fragment compiled from `doc` (shared/blocks.ts markup, inline styles). */
  html: string;
}

export interface EmailEditorProps {
  /**
   * INITIAL document (TipTap JSON). Pass `null` for empty. When only HTML exists (legacy
   * template), pass `initialHtml`. Read on mount and whenever `contentKey` changes; later
   * changes to `doc`/`initialHtml` alone are ignored, so feeding onChange results back in
   * never resets the cursor or undo history.
   */
  doc: unknown | null;
  initialHtml?: string;
  onChange: (value: EmailEditorValue) => void;
  /** Theme whose body CSS styles the editing canvas (WYSIWYG with the sent email). Changing it only restyles. */
  theme: ThemeId;
  /** Placeholder names offered in the "Insert field" menu, e.g. ['Name', 'Email', 'Department']. */
  placeholders?: string[];
  /** 'body' (full toolbar) or 'signature' (compact toolbar, no custom blocks). */
  variant?: 'body' | 'signature';
  minHeight?: number;
  className?: string;
  /** Change this (e.g. to a template id + load counter) to reload content from `doc` / `initialHtml`. Clears undo history. */
  contentKey?: string | number;
  /** True while pasted/dropped images are still uploading (block saving/sending meanwhile). */
  onBusyChange?: (busy: boolean) => void;
  /** Read-only view (e.g. shared templates): hides the toolbar. */
  readOnly?: boolean;
  /** Receives the TipTap editor instance (null on unmount), for advanced integrations. */
  onEditor?: (editor: Editor | null) => void;
  /** Accessible name for the editing area. */
  ariaLabel?: string;
}

const ENHANCED_FONTS_URL = `https://fonts.googleapis.com/css2?${FONT_CATALOG.filter(f => f.googleFamily).map(f => `family=${f.googleFamily}`).join('&')}&display=swap`;

function ensureStylesheet(id: string, href: string) {
  if (typeof document === 'undefined' || !href || document.getElementById(id)) return;
  const link = document.createElement('link');
  link.id = id;
  link.rel = 'stylesheet';
  link.href = href;
  document.head.appendChild(link);
}

function initialContent(doc: unknown, html: string | undefined): JSONContent | string {
  if (doc && typeof doc === 'object' && (doc as JSONContent).type === 'doc') return doc as JSONContent;
  if (html && html.trim()) return prepareHtmlForEditor(html);
  return '';
}

export function EmailEditor(props: EmailEditorProps) {
  const {
    doc, initialHtml, theme, placeholders = [], variant = 'body', minHeight = 320, className,
    contentKey, readOnly = false, ariaLabel,
  } = props;

  // Latest callbacks without re-creating the editor.
  const onChangeRef = useRef(props.onChange);
  const onBusyRef = useRef(props.onBusyChange);
  const onEditorRef = useRef(props.onEditor);
  onChangeRef.current = props.onChange;
  onBusyRef.current = props.onBusyChange;
  onEditorRef.current = props.onEditor;
  const editorRef = useRef<Editor | null>(null);
  const openLinkRef = useRef<() => void>(() => {});
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const rawId = useId();
  const scopeId = rawId.replace(/[^a-zA-Z0-9_-]/g, '');

  useEffect(() => {
    ensureStylesheet('kb-editor-enhanced-fonts', ENHANCED_FONTS_URL);
  }, []);
  useEffect(() => {
    ensureStylesheet(`kb-theme-fonts-${theme}`, themeWebfontUrl(theme));
  }, [theme]);

  const notifyError = useCallback((m: string) => toast.error(m), []);

  const extensions = useMemo(() => {
    const shortcuts = Extension.create({
      name: 'kbEditorShortcuts',
      addKeyboardShortcuts() {
        return { 'Mod-k': () => { openLinkRef.current(); return true; } };
      },
    });
    return [
      ...buildExtensions({
        variant,
        uploads: {
          uploadFile: (blob, filename) => uploadFile(blob, 'image', { filename }),
          rehostImage,
          onError: m => notifyError(m),
          onBusyChange: b => onBusyRef.current?.(b),
        },
      }),
      shortcuts,
    ];
  }, [variant, notifyError]);

  const emit = useCallback(() => {
    if (timer.current) { clearTimeout(timer.current); timer.current = null; }
    const ed = editorRef.current;
    if (!ed || ed.isDestroyed) return;
    const json = ed.getJSON();
    onChangeRef.current({ doc: json, html: serializeEmail(json) });
  }, []);

  const editor = useEditor({
    extensions,
    content: initialContent(doc, initialHtml),
    editable: !readOnly,
    immediatelyRender: true,
    shouldRerenderOnTransaction: false,
    editorProps: {
      attributes: {
        class: cn('kb-prose', variant === 'signature' ? 'kb-signature' : 'kb-body'),
        'aria-label': ariaLabel ?? (variant === 'signature' ? 'Signature' : 'Message body'),
        role: 'textbox',
        'aria-multiline': 'true',
        spellcheck: 'true',
      },
      transformPastedHTML: html => {
        try { return normalizePastedHtml(html); } catch { return html; }
      },
      clipboardTextParser: (text, $context, _plain, view) => {
        const div = document.createElement('div');
        div.innerHTML = plainTextToHtml(text);
        const slice = PMDOMParser.fromSchema(view.state.schema).parseSlice(div, { preserveWhitespace: true, context: $context });
        return Slice.maxOpen(slice.content);
      },
      handlePaste: (_view, event) => {
        const dt = event.clipboardData;
        if (!dt || !editorRef.current) return false;
        const files = Array.from(dt.files || []).filter(f => f.type.startsWith('image/'));
        if (files.length && !dt.getData('text/html')) {
          insertImageFiles(editorRef.current, files, undefined, notifyError);
          return true;
        }
        return false;
      },
      handleDrop: (view, event, _slice, moved) => {
        if (moved || !editorRef.current) return false;
        const files = Array.from(event.dataTransfer?.files || []).filter(f => f.type.startsWith('image/'));
        if (!files.length) return false;
        event.preventDefault();
        const pos = view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos;
        insertImageFiles(editorRef.current, files, pos, notifyError);
        return true;
      },
    },
    onUpdate: () => {
      if (timer.current) clearTimeout(timer.current);
      timer.current = setTimeout(emit, 250);
    },
    onBlur: () => { if (timer.current) emit(); },
  }, [extensions]);

  editorRef.current = editor;

  useEffect(() => {
    onEditorRef.current?.(editor);
    return () => {
      // Flush a pending change so nothing typed in the last 250 ms is lost.
      if (timer.current) emit();
      onEditorRef.current?.(null);
    };
  }, [editor, emit]);

  // Reload only on a deliberate external change (contentKey), never because `doc` identity changed.
  const firstKey = useRef(true);
  const latestInitial = useRef({ doc, initialHtml });
  latestInitial.current = { doc, initialHtml };
  useEffect(() => {
    if (firstKey.current) { firstKey.current = false; return; }
    if (!editor || editor.isDestroyed) return;
    if (timer.current) { clearTimeout(timer.current); timer.current = null; }
    const { doc: d, initialHtml: h } = latestInitial.current;
    editor.commands.setContent(initialContent(d, h), { emitUpdate: false });
    // Fresh state so undo cannot step back into the previous template.
    editor.view.updateState(EditorState.create({ doc: editor.state.doc, plugins: editor.state.plugins }));
  }, [contentKey, editor]);

  useEffect(() => {
    if (editor && !editor.isDestroyed && editor.isEditable === readOnly) editor.setEditable(!readOnly);
  }, [editor, readOnly]);

  // ---------------------------------------------------------------- dialogs
  const [link, setLink] = useState<{ open: boolean; href: string; text: string; showText: boolean; editing: boolean }>({ open: false, href: '', text: '', showText: true, editing: false });
  const [imageOpen, setImageOpen] = useState(false);
  const [buttonOpen, setButtonOpen] = useState(false);

  const openLink = useCallback(() => {
    const ed = editorRef.current;
    if (!ed || !ed.isEditable) return;
    const editing = ed.isActive('link');
    if (editing) ed.chain().extendMarkRange('link').run();
    const { from, to, empty } = ed.state.selection;
    const text = empty ? '' : ed.state.doc.textBetween(from, to, ' ');
    const sameBlock = ed.state.selection.$from.sameParent(ed.state.selection.$to);
    setLink({ open: true, href: editing ? String(ed.getAttributes('link').href || '') : '', text, showText: empty || sameBlock, editing });
  }, []);
  openLinkRef.current = openLink;

  const submitLink = ({ href, text }: { href: string; text: string }) => {
    const ed = editorRef.current;
    if (!ed) return;
    const { empty } = ed.state.selection;
    const original = link.text;
    const c = ed.chain().focus();
    if (empty) {
      const label = text || href.replace(/^mailto:/, '');
      c.insertContent({ type: 'text', text: label, marks: [{ type: 'link', attrs: { href } }] }).run();
    } else if (link.showText && text && text !== original) {
      const marks = ed.state.selection.$from.marks().filter(m => m.type.name !== 'link').map(m => m.toJSON());
      c.insertContent({ type: 'text', text, marks: [...marks, { type: 'link', attrs: { href } }] }).run();
    } else {
      c.extendMarkRange('link').setLink({ href }).run();
    }
    setLink(l => ({ ...l, open: false }));
  };

  const removeLink = () => {
    editorRef.current?.chain().focus().extendMarkRange('link').unsetLink().run();
    setLink(l => ({ ...l, open: false }));
  };

  const emptyButton: ButtonAttrs = { label: '', href: 'https://', variant: 'primary', align: 'left' };

  if (!editor) return null;

  const scope = `[data-kb-editor="${scopeId}"] .kb-body`;
  const paperPad = variant === 'signature' ? 24 : BODY_PADDING_X;

  return (
    <div className={cn('kb-editor flex flex-col overflow-hidden rounded-lg border border-border bg-card', className)} data-kb-editor={scopeId} data-variant={variant}>
      <style>{themeCss(theme, scope)}</style>
      {!readOnly && (
        <EditorToolbar
          editor={editor}
          variant={variant}
          placeholders={placeholders}
          onLink={openLink}
          onImage={() => setImageOpen(true)}
          onButton={() => setButtonOpen(true)}
        />
      )}
      <div className="kb-canvas" onClick={e => { if (e.target === e.currentTarget) editor.chain().focus('end').run(); }}>
        <div
          className={cn('kb-paper', variant === 'signature' ? 'kb-body' : undefined)}
          data-variant={variant}
          style={{ maxWidth: EMAIL_WIDTH, minHeight, ['--kb-paper-pad' as string]: `${paperPad}px` }}
        >
          <EditorContent editor={editor} />
        </div>
      </div>
      <LinkDialog
        open={link.open}
        onOpenChange={o => setLink(l => ({ ...l, open: o }))}
        initialHref={link.href}
        initialText={link.text}
        showText={link.showText}
        isEditing={link.editing}
        onSubmit={submitLink}
        onRemove={removeLink}
      />
      <ImageDialog
        open={imageOpen}
        onOpenChange={setImageOpen}
        onInsert={v => editor.chain().focus().insertContent({ type: 'image', attrs: { src: v.src, alt: v.alt, width: v.width } }).run()}
      />
      <ButtonDialog
        open={buttonOpen}
        onOpenChange={setButtonOpen}
        initial={emptyButton}
        onSubmit={a => { editor.chain().focus().insertKbButton(a).run(); setButtonOpen(false); }}
      />
    </div>
  );
}
