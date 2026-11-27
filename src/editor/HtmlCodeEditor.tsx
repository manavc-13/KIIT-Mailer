// <HtmlCodeEditor>: CodeMirror HTML source editor (html mode) with Format, Insert field
// and a size counter (Gmail clips messages over ~102 KB).
import { useEffect, useMemo, useRef, useState } from 'react';
import CodeMirror, { EditorView, type ReactCodeMirrorRef } from '@uiw/react-codemirror';
import { html as htmlLang } from '@codemirror/lang-html';
import beautify from 'js-beautify';
import { Braces, ChevronDown, WandSparkles } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuLabel, DropdownMenuSeparator, DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import { cn } from '@/lib/utils';

export interface HtmlCodeEditorProps {
  value: string;
  onChange: (html: string) => void;
  placeholders?: string[];
  minHeight?: number;
  className?: string;
  readOnly?: boolean;
}

const GMAIL_CLIP_BYTES = 102 * 1024;

/** Pretty-print HTML (2-space indent), keeping inline tags inline. */
export function formatHtml(src: string): string {
  return beautify.html(src, {
    indent_size: 2,
    wrap_line_length: 0,
    preserve_newlines: true,
    max_preserve_newlines: 1,
    indent_inner_html: true,
    end_with_newline: false,
    extra_liners: [],
    inline: ['a', 'abbr', 'b', 'br', 'code', 'em', 'i', 'img', 'small', 'span', 'strong', 'sub', 'sup', 'u', 's', 'font'],
  });
}

function useDarkMode(): boolean {
  const get = () => typeof document !== 'undefined' && document.documentElement.classList.contains('dark');
  const [dark, setDark] = useState(get);
  useEffect(() => {
    const obs = new MutationObserver(() => setDark(get()));
    obs.observe(document.documentElement, { attributes: true, attributeFilter: ['class'] });
    return () => obs.disconnect();
  }, []);
  return dark;
}

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  return `${(n / 1024).toFixed(n < 10 * 1024 ? 1 : 0)} KB`;
}

export function HtmlCodeEditor({ value, onChange, placeholders = [], minHeight = 320, className, readOnly = false }: HtmlCodeEditorProps) {
  const ref = useRef<ReactCodeMirrorRef>(null);
  const dark = useDarkMode();
  const [custom, setCustom] = useState('');
  const extensions = useMemo(() => [htmlLang({ selfClosingTags: false }), EditorView.lineWrapping], []);
  const bytes = useMemo(() => new TextEncoder().encode(value).length, [value]);
  const clipped = bytes > GMAIL_CLIP_BYTES;

  const insert = (text: string) => {
    const view = ref.current?.view;
    if (!view) { onChange(value + text); return; }
    const { from, to } = view.state.selection.main;
    view.dispatch({ changes: { from, to, insert: text }, selection: { anchor: from + text.length }, scrollIntoView: true });
    view.focus();
  };

  const format = () => {
    try { onChange(formatHtml(value)); } catch { /* leave as is */ }
  };

  return (
    <div className={cn('flex flex-col overflow-hidden rounded-lg border border-border bg-card', className)}>
      <div className="flex flex-wrap items-center gap-1 border-b border-border px-2 py-1.5" role="toolbar" aria-label="HTML tools">
        <Tooltip>
          <TooltipTrigger asChild>
            <Button type="button" size="sm" variant="ghost" onClick={format} disabled={readOnly || !value.trim()} aria-label="Format HTML">
              <WandSparkles aria-hidden /> Format
            </Button>
          </TooltipTrigger>
          <TooltipContent side="bottom">Re-indent the HTML (2 spaces)</TooltipContent>
        </Tooltip>
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button type="button" size="sm" variant="ghost" disabled={readOnly} aria-label="Insert recipient field">
              <Braces aria-hidden /> Insert field <ChevronDown className="opacity-60" aria-hidden />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start" className="w-60">
            <DropdownMenuLabel>Recipient fields</DropdownMenuLabel>
            {placeholders.length === 0 && <p className="px-2 pb-1 text-xs text-muted-foreground">Add recipients to list their columns here, or type a field name below.</p>}
            {placeholders.map(p => (
              <DropdownMenuItem key={p} onSelect={() => insert(`{${p}}`)}><span className="kb-field-preview">{`{${p}}`}</span></DropdownMenuItem>
            ))}
            <DropdownMenuSeparator />
            <form className="flex gap-1 p-1" onSubmit={e => { e.preventDefault(); if (custom.trim()) { insert(`{${custom.trim()}}`); setCustom(''); } }} onKeyDown={e => e.stopPropagation()}>
              <Input value={custom} onChange={e => setCustom(e.target.value.replace(/[{}<>]/g, ''))} placeholder="Field name" aria-label="Custom field name" className="h-8 text-xs" />
              <Button type="submit" size="sm" variant="outline" disabled={!custom.trim()}>Add</Button>
            </form>
          </DropdownMenuContent>
        </DropdownMenu>
        <span className="ml-auto flex items-center gap-2 text-xs tabular-nums text-muted-foreground" aria-live="polite">
          <span>{value.length.toLocaleString()} characters</span>
          <span aria-hidden>·</span>
          <span className={cn(clipped && 'font-medium text-warning')} title={clipped ? 'Gmail clips messages larger than about 102 KB and hides the rest behind “View entire message”.' : 'Gmail clips messages larger than about 102 KB.'}>
            {formatBytes(bytes)}{clipped ? ' · over Gmail’s 102 KB clip limit' : ''}
          </span>
        </span>
      </div>
      <CodeMirror
        ref={ref}
        value={value}
        onChange={onChange}
        extensions={extensions}
        theme={dark ? 'dark' : 'light'}
        readOnly={readOnly}
        minHeight={`${minHeight}px`}
        basicSetup={{ lineNumbers: true, foldGutter: true, highlightActiveLine: true, autocompletion: true, bracketMatching: true, closeBrackets: true }}
        aria-label="HTML source"
        className="kb-code text-[13px]"
      />
    </div>
  );
}
