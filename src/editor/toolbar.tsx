// Sticky, grouped formatting toolbar for the email editor.
import { useState, type ReactNode } from 'react';
import { useEditorState, type Editor } from '@tiptap/react';
import {
  AlignCenter, AlignJustify, AlignLeft, AlignRight, Baseline, Bold, Braces, ChevronDown, Highlighter, ImagePlus, IndentDecrease,
  IndentIncrease, Italic, Link2, List, ListOrdered, Minus, MoveVertical, Plus, Redo2, RemoveFormatting, SeparatorHorizontal,
  SquarePlus, Strikethrough, Subscript as SubIcon, Superscript as SupIcon, Table2, Underline as UnderlineIcon, Undo2,
} from 'lucide-react';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover';
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuLabel, DropdownMenuSeparator, DropdownMenuSub,
  DropdownMenuSubContent, DropdownMenuSubTrigger, DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import { Select, SelectContent, SelectGroup, SelectItem, SelectLabel, SelectSeparator, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Input } from '@/components/ui/input';
import { Button } from '@/components/ui/button';
import { cn } from '@/lib/utils';
import { FONT_CATALOG, FONT_SIZES_PX, LINE_HEIGHTS, findFont } from '../../shared/fonts';
import { HIGHLIGHT_COLORS, TEXT_COLORS, CALLOUT_VARIANTS, type Swatch } from './constants';
import { normalizeColor } from './style-utils';

const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
const mod = isMac ? '⌘' : 'Ctrl';

// ------------------------------------------------------------------ primitives

function Tip({ label, shortcut, children }: { label: string; shortcut?: string; children: ReactNode }) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>{children}</TooltipTrigger>
      <TooltipContent side="bottom">
        {label}{shortcut && <span className="ml-2 opacity-70">{shortcut}</span>}
      </TooltipContent>
    </Tooltip>
  );
}

const tbBtn = 'inline-flex h-8 min-w-8 items-center justify-center gap-1 rounded-md px-1.5 text-sm text-foreground transition-colors hover:bg-muted focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/60 disabled:pointer-events-none disabled:opacity-40 aria-pressed:bg-accent aria-pressed:text-accent-foreground data-[state=open]:bg-muted [&_svg]:size-4 [&_svg]:shrink-0';

function ToolButton({ label, shortcut, active, disabled, onClick, children }: {
  label: string; shortcut?: string; active?: boolean; disabled?: boolean; onClick: () => void; children: ReactNode;
}) {
  return (
    <Tip label={label} shortcut={shortcut}>
      <button
        type="button"
        className={tbBtn}
        aria-label={label}
        aria-pressed={active === undefined ? undefined : active}
        disabled={disabled}
        onMouseDown={e => e.preventDefault()}
        onClick={onClick}
      >
        {children}
      </button>
    </Tip>
  );
}

function Group({ children, label }: { children: ReactNode; label: string }) {
  return <div role="group" aria-label={label} className="flex items-center gap-0.5">{children}</div>;
}

const Sep = () => <span className="mx-1 h-5 w-px shrink-0 bg-border" aria-hidden />;

// ------------------------------------------------------------------ colour picker

function ColorPicker({ label, icon, swatches, value, onPick, onClear, clearLabel }: {
  label: string; icon: ReactNode; swatches: Swatch[]; value: string | null; onPick: (c: string) => void; onClear: () => void; clearLabel: string;
}) {
  const [open, setOpen] = useState(false);
  const [hex, setHex] = useState('');
  const apply = (c: string) => { onPick(c); setOpen(false); };
  const hexOk = normalizeColor(hex.startsWith('#') ? hex : `#${hex}`);
  return (
    <Popover open={open} onOpenChange={o => { setOpen(o); if (o) setHex(value || ''); }}>
      <Tip label={label}>
        <PopoverTrigger asChild>
          <button type="button" className={tbBtn} aria-label={label} onMouseDown={e => e.preventDefault()}>
            <span className="relative inline-flex flex-col items-center">
              {icon}
              <span className="mt-px h-[3px] w-4 rounded-sm border border-border/60" style={{ backgroundColor: value || 'transparent' }} aria-hidden />
            </span>
            <ChevronDown className="!size-3 opacity-60" aria-hidden />
          </button>
        </PopoverTrigger>
      </Tip>
      <PopoverContent className="w-60 p-3" align="start" onOpenAutoFocus={e => e.preventDefault()}>
        <p className="mb-2 text-xs font-medium text-muted-foreground">{label}</p>
        <div className="grid grid-cols-5 gap-1.5" role="listbox" aria-label={`${label} palette`}>
          {swatches.map(s => (
            <button
              key={s.value}
              type="button"
              role="option"
              aria-selected={value === s.value}
              aria-label={s.label}
              title={`${s.label} ${s.value}`}
              onClick={() => apply(s.value)}
              className={cn('size-8 rounded border border-border/80 outline-offset-1 hover:scale-105 focus-visible:outline-2 focus-visible:outline-ring', value === s.value && 'outline-2 outline-ring')}
              style={{ backgroundColor: s.value }}
            />
          ))}
        </div>
        <form className="mt-3 flex items-center gap-1.5" onSubmit={e => { e.preventDefault(); if (hexOk) apply(hexOk); }}>
          <span className="size-8 shrink-0 rounded border border-border" style={{ backgroundColor: hexOk || 'transparent' }} aria-hidden />
          <Input value={hex} onChange={e => setHex(e.target.value.trim())} placeholder="#1f3864" aria-label="Hex colour" className="h-8 font-mono text-xs" maxLength={7} />
          <Button type="submit" size="sm" variant="outline" disabled={!hexOk}>Apply</Button>
        </form>
        <Button type="button" variant="ghost" size="sm" className="mt-2 w-full justify-start" onClick={() => { onClear(); setOpen(false); }}>{clearLabel}</Button>
      </PopoverContent>
    </Popover>
  );
}

// ------------------------------------------------------------------ font controls

function FontFamilySelect({ editor, value }: { editor: Editor; value: string | null }) {
  const font = findFont(value);
  const current = value ? font?.id ?? 'custom' : 'default';
  const universal = FONT_CATALOG.filter(f => f.group === 'universal');
  const enhanced = FONT_CATALOG.filter(f => f.group === 'enhanced');
  const firstName = (stack: string) => stack.split(',')[0].replace(/['"]/g, '').trim();
  return (
    <Select
      value={current}
      onValueChange={id => {
        const f = FONT_CATALOG.find(x => x.id === id);
        if (id === 'default' || !f) editor.chain().focus().unsetFontFamily().run();
        else editor.chain().focus().setFontFamily(f.stack).run();
      }}
    >
      <Tip label="Font">
        <SelectTrigger size="sm" className="h-8 w-[9.5rem] border-transparent px-2 shadow-none hover:bg-muted" aria-label="Font family">
          <SelectValue placeholder="Font">
            <span className="truncate" style={{ fontFamily: font?.stack }}>{current === 'default' ? 'Theme font' : font?.label ?? firstName(value || '')}</span>
          </SelectValue>
        </SelectTrigger>
      </Tip>
      <SelectContent position="popper" className="max-h-96 w-80">
        <SelectItem value="default">Theme font (default)</SelectItem>
        <SelectSeparator />
        <SelectGroup>
          <SelectLabel>Universal · renders everywhere</SelectLabel>
          {universal.map(f => (
            <SelectItem key={f.id} value={f.id}><span style={{ fontFamily: f.stack }}>{f.label}</span></SelectItem>
          ))}
        </SelectGroup>
        <SelectSeparator />
        <SelectGroup>
          <SelectLabel className="leading-snug">Enhanced · Apple Mail, iOS; falls back in Gmail/Outlook</SelectLabel>
          {enhanced.map(f => (
            <SelectItem key={f.id} value={f.id}>
              <span className="flex w-full items-baseline justify-between gap-3">
                <span style={{ fontFamily: f.stack }}>{f.label}</span>
                <span className="text-xs text-muted-foreground" style={{ fontFamily: f.fallback }}>→ {firstName(f.fallback)}</span>
              </span>
            </SelectItem>
          ))}
        </SelectGroup>
      </SelectContent>
    </Select>
  );
}

function FontSizePicker({ editor, value }: { editor: Editor; value: string | null }) {
  const [open, setOpen] = useState(false);
  const [custom, setCustom] = useState('');
  const px = value ? parseInt(value, 10) : null;
  const apply = (n: number | null) => {
    if (n === null) editor.chain().focus().unsetFontSize().run();
    else editor.chain().focus().setFontSize(`${Math.max(8, Math.min(72, Math.round(n)))}px`).run();
    setOpen(false);
  };
  return (
    <Popover open={open} onOpenChange={o => { setOpen(o); if (o) setCustom(px ? String(px) : ''); }}>
      <Tip label="Font size">
        <PopoverTrigger asChild>
          <button type="button" className={cn(tbBtn, 'w-14 justify-between tabular-nums')} aria-label="Font size" onMouseDown={e => e.preventDefault()}>
            <span className="text-xs">{px ?? 'Auto'}</span>
            <ChevronDown className="!size-3 opacity-60" aria-hidden />
          </button>
        </PopoverTrigger>
      </Tip>
      <PopoverContent className="w-44 p-2" align="start" onOpenAutoFocus={e => e.preventDefault()}>
        <form className="mb-2 flex gap-1" onSubmit={e => { e.preventDefault(); const n = parseFloat(custom); if (n > 0) apply(n); }}>
          <Input value={custom} onChange={e => setCustom(e.target.value.replace(/[^\d.]/g, ''))} placeholder="px" aria-label="Custom font size in pixels" className="h-8 text-xs" inputMode="numeric" />
          <Button type="submit" size="sm" variant="outline">Set</Button>
        </form>
        <div className="max-h-60 overflow-y-auto" role="listbox" aria-label="Font sizes">
          <button type="button" role="option" aria-selected={px === null} className="w-full rounded px-2 py-1 text-left text-sm hover:bg-muted" onClick={() => apply(null)}>Theme default</button>
          {FONT_SIZES_PX.map(s => (
            <button key={s} type="button" role="option" aria-selected={px === s} className={cn('w-full rounded px-2 py-1 text-left text-sm tabular-nums hover:bg-muted', px === s && 'bg-accent text-accent-foreground')} onClick={() => apply(s)}>{s}px</button>
          ))}
        </div>
      </PopoverContent>
    </Popover>
  );
}

// ------------------------------------------------------------------ toolbar

export interface ToolbarProps {
  editor: Editor;
  variant: 'body' | 'signature';
  placeholders: string[];
  onLink: () => void;
  onImage: () => void;
  onButton: () => void;
}

export function EditorToolbar({ editor, variant, placeholders, onLink, onImage, onButton }: ToolbarProps) {
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const ts = e.getAttributes('textStyle');
      const block = e.isActive('heading', { level: 1 }) ? 'h1' : e.isActive('heading', { level: 2 }) ? 'h2' : e.isActive('heading', { level: 3 }) ? 'h3' : 'p';
      const para = e.isActive('heading') ? e.getAttributes('heading') : e.getAttributes('paragraph');
      const cell = e.isActive('tableHeader') ? e.getAttributes('tableHeader') : e.getAttributes('tableCell');
      return {
        canUndo: e.can().undo(),
        canRedo: e.can().redo(),
        block,
        fontFamily: (ts.fontFamily as string) || null,
        fontSize: (ts.fontSize as string) || null,
        color: normalizeColor(ts.color) || null,
        bg: normalizeColor(ts.backgroundColor) || null,
        bold: e.isActive('bold'),
        italic: e.isActive('italic'),
        underline: e.isActive('underline'),
        strike: e.isActive('strike'),
        sup: e.isActive('superscript'),
        sub: e.isActive('subscript'),
        align: (['center', 'right', 'justify'].find(a => e.isActive({ textAlign: a })) || 'left') as string,
        lineHeight: (para.lineHeight as string) || null,
        bullet: e.isActive('bulletList'),
        ordered: e.isActive('orderedList'),
        link: e.isActive('link'),
        inTable: e.isActive('table'),
        cellBg: (cell.backgroundColor as string) || null,
      };
    },
  });
  const chain = () => editor.chain().focus();
  const full = variant === 'body';

  return (
    <div className="kb-toolbar sticky top-0 z-20 flex flex-wrap items-center gap-y-1 border-b border-border bg-card px-2 py-1.5" role="toolbar" aria-label="Formatting" aria-orientation="horizontal">
      <Group label="History">
        <ToolButton label="Undo" shortcut={`${mod}+Z`} disabled={!s.canUndo} onClick={() => chain().undo().run()}><Undo2 /></ToolButton>
        <ToolButton label="Redo" shortcut={`${mod}+Shift+Z`} disabled={!s.canRedo} onClick={() => chain().redo().run()}><Redo2 /></ToolButton>
      </Group>
      <Sep />
      {full && (
        <>
          <Select
            value={s.block}
            onValueChange={v => {
              if (v === 'p') chain().setParagraph().run();
              else chain().setHeading({ level: Number(v.slice(1)) as 1 | 2 | 3 }).run();
            }}
          >
            <Tip label="Text style">
              <SelectTrigger size="sm" className="h-8 w-[7.5rem] border-transparent px-2 shadow-none hover:bg-muted" aria-label="Text style"><SelectValue /></SelectTrigger>
            </Tip>
            <SelectContent position="popper">
              <SelectItem value="p">Paragraph</SelectItem>
              <SelectItem value="h1"><span className="text-lg font-semibold">Heading 1</span></SelectItem>
              <SelectItem value="h2"><span className="text-base font-semibold">Heading 2</span></SelectItem>
              <SelectItem value="h3"><span className="text-sm font-semibold">Heading 3</span></SelectItem>
            </SelectContent>
          </Select>
          <Sep />
        </>
      )}
      <Group label="Font">
        <FontFamilySelect editor={editor} value={s.fontFamily} />
        <FontSizePicker editor={editor} value={s.fontSize} />
      </Group>
      <Sep />
      <Group label="Text formatting">
        <ToolButton label="Bold" shortcut={`${mod}+B`} active={s.bold} onClick={() => chain().toggleBold().run()}><Bold /></ToolButton>
        <ToolButton label="Italic" shortcut={`${mod}+I`} active={s.italic} onClick={() => chain().toggleItalic().run()}><Italic /></ToolButton>
        <ToolButton label="Underline" shortcut={`${mod}+U`} active={s.underline} onClick={() => chain().toggleUnderline().run()}><UnderlineIcon /></ToolButton>
        <ToolButton label="Strikethrough" shortcut={`${mod}+Shift+S`} active={s.strike} onClick={() => chain().toggleStrike().run()}><Strikethrough /></ToolButton>
        {full && <ToolButton label="Superscript" shortcut={`${mod}+.`} active={s.sup} onClick={() => chain().toggleSuperscript().run()}><SupIcon /></ToolButton>}
        {full && <ToolButton label="Subscript" shortcut={`${mod}+,`} active={s.sub} onClick={() => chain().toggleSubscript().run()}><SubIcon /></ToolButton>}
      </Group>
      <Sep />
      <Group label="Colour">
        <ColorPicker label="Text colour" icon={<Baseline />} swatches={TEXT_COLORS} value={s.color} onPick={c => chain().setColor(c).run()} onClear={() => chain().unsetColor().run()} clearLabel="Theme colour (default)" />
        <ColorPicker label="Highlight" icon={<Highlighter />} swatches={HIGHLIGHT_COLORS} value={s.bg} onPick={c => chain().setBackgroundColor(c).run()} onClear={() => chain().unsetBackgroundColor().run()} clearLabel="No highlight" />
      </Group>
      <Sep />
      <Group label="Paragraph">
        <ToolButton label="Align left" shortcut={`${mod}+Shift+L`} active={s.align === 'left'} onClick={() => chain().unsetTextAlign().run()}><AlignLeft /></ToolButton>
        <ToolButton label="Align centre" shortcut={`${mod}+Shift+E`} active={s.align === 'center'} onClick={() => chain().setTextAlign('center').run()}><AlignCenter /></ToolButton>
        <ToolButton label="Align right" shortcut={`${mod}+Shift+R`} active={s.align === 'right'} onClick={() => chain().setTextAlign('right').run()}><AlignRight /></ToolButton>
        {full && <ToolButton label="Justify" shortcut={`${mod}+Shift+J`} active={s.align === 'justify'} onClick={() => chain().setTextAlign('justify').run()}><AlignJustify /></ToolButton>}
        {full && (
          <DropdownMenu>
            <Tip label="Line spacing">
              <DropdownMenuTrigger asChild>
                <button type="button" className={tbBtn} aria-label="Line spacing" onMouseDown={e => e.preventDefault()}><MoveVertical /><ChevronDown className="!size-3 opacity-60" aria-hidden /></button>
              </DropdownMenuTrigger>
            </Tip>
            <DropdownMenuContent align="start" onCloseAutoFocus={e => e.preventDefault()}>
              <DropdownMenuLabel>Line spacing</DropdownMenuLabel>
              <DropdownMenuItem onSelect={() => chain().setBlockLineHeight(null).run()}>
                <span className={cn(!s.lineHeight && 'font-semibold')}>Theme default</span>
              </DropdownMenuItem>
              {LINE_HEIGHTS.map(lh => (
                <DropdownMenuItem key={lh} onSelect={() => chain().setBlockLineHeight(lh).run()}>
                  <span className={cn('tabular-nums', s.lineHeight === lh && 'font-semibold')}>{lh}</span>
                </DropdownMenuItem>
              ))}
            </DropdownMenuContent>
          </DropdownMenu>
        )}
      </Group>
      <Sep />
      <Group label="Lists and indentation">
        <ToolButton label="Bulleted list" shortcut={`${mod}+Shift+8`} active={s.bullet} onClick={() => chain().toggleBulletList().run()}><List /></ToolButton>
        <ToolButton label="Numbered list" shortcut={`${mod}+Shift+7`} active={s.ordered} onClick={() => chain().toggleOrderedList().run()}><ListOrdered /></ToolButton>
        {full && <ToolButton label="Decrease indent" shortcut="Shift+Tab" onClick={() => chain().outdentBlock().run()}><IndentDecrease /></ToolButton>}
        {full && <ToolButton label="Increase indent" shortcut="Tab" onClick={() => chain().indentBlock().run()}><IndentIncrease /></ToolButton>}
      </Group>
      <Sep />
      <Group label="Insert">
        <ToolButton label="Link" shortcut={`${mod}+K`} active={s.link} onClick={onLink}><Link2 /></ToolButton>
        <ToolButton label="Image" onClick={onImage}><ImagePlus /></ToolButton>
        {full && (
          <DropdownMenu>
            <Tip label="Table">
              <DropdownMenuTrigger asChild>
                <button type="button" className={tbBtn} aria-label="Table" aria-pressed={s.inTable} onMouseDown={e => e.preventDefault()}><Table2 /><ChevronDown className="!size-3 opacity-60" aria-hidden /></button>
              </DropdownMenuTrigger>
            </Tip>
            <DropdownMenuContent align="start" className="w-56" onCloseAutoFocus={e => e.preventDefault()}>
              <DropdownMenuLabel>Insert table</DropdownMenuLabel>
              <DropdownMenuItem onSelect={() => chain().insertTable({ rows: 2, cols: 2, withHeaderRow: false }).run()}>2 × 2</DropdownMenuItem>
              <DropdownMenuItem onSelect={() => chain().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()}>3 × 3 with header row</DropdownMenuItem>
              <DropdownMenuItem onSelect={() => chain().insertTable({ rows: 4, cols: 2, withHeaderRow: true }).run()}>4 × 2 with header row</DropdownMenuItem>
              {s.inTable && (
                <>
                  <DropdownMenuSeparator />
                  <DropdownMenuLabel>Rows and columns</DropdownMenuLabel>
                  <DropdownMenuItem onSelect={() => chain().addRowBefore().run()}><Plus /> Row above</DropdownMenuItem>
                  <DropdownMenuItem onSelect={() => chain().addRowAfter().run()}><Plus /> Row below</DropdownMenuItem>
                  <DropdownMenuItem onSelect={() => chain().addColumnBefore().run()}><Plus /> Column left</DropdownMenuItem>
                  <DropdownMenuItem onSelect={() => chain().addColumnAfter().run()}><Plus /> Column right</DropdownMenuItem>
                  <DropdownMenuItem onSelect={() => chain().deleteRow().run()}><Minus /> Delete row</DropdownMenuItem>
                  <DropdownMenuItem onSelect={() => chain().deleteColumn().run()}><Minus /> Delete column</DropdownMenuItem>
                  <DropdownMenuItem onSelect={() => chain().toggleHeaderRow().run()}>Toggle header row</DropdownMenuItem>
                  <DropdownMenuSub>
                    <DropdownMenuSubTrigger>Cell background</DropdownMenuSubTrigger>
                    <DropdownMenuSubContent className="w-48 p-2">
                      <div className="grid grid-cols-5 gap-1.5">
                        {HIGHLIGHT_COLORS.filter(c => c.value !== '#0c1a28' && c.value !== '#b8892f').concat([{ label: 'White', value: '#ffffff' }]).map(c => (
                          <DropdownMenuItem key={c.value} className="size-7 p-0" aria-label={c.label} title={c.label} onSelect={() => chain().setCellAttribute('backgroundColor', c.value).run()}>
                            <span className="size-full rounded border border-border" style={{ backgroundColor: c.value }} />
                          </DropdownMenuItem>
                        ))}
                      </div>
                      <DropdownMenuItem className="mt-1" onSelect={() => chain().setCellAttribute('backgroundColor', null).run()}>No background</DropdownMenuItem>
                    </DropdownMenuSubContent>
                  </DropdownMenuSub>
                  <DropdownMenuSeparator />
                  <DropdownMenuItem variant="destructive" onSelect={() => chain().deleteTable().run()}>Delete table</DropdownMenuItem>
                </>
              )}
            </DropdownMenuContent>
          </DropdownMenu>
        )}
        {full && <ToolButton label="Horizontal line" onClick={() => chain().setHorizontalRule().run()}><SeparatorHorizontal /></ToolButton>}
        {full && (
          <DropdownMenu>
            <Tip label="Insert block">
              <DropdownMenuTrigger asChild>
                <button type="button" className={tbBtn} aria-label="Insert block" onMouseDown={e => e.preventDefault()}><SquarePlus /><span className="hidden text-xs xl:inline">Block</span><ChevronDown className="!size-3 opacity-60" aria-hidden /></button>
              </DropdownMenuTrigger>
            </Tip>
            <DropdownMenuContent align="start" className="w-60" onCloseAutoFocus={e => e.preventDefault()}>
              <DropdownMenuLabel>Email blocks</DropdownMenuLabel>
              <DropdownMenuItem onSelect={onButton}>Button…<span className="ml-auto text-xs text-muted-foreground">call to action</span></DropdownMenuItem>
              <DropdownMenuSub>
                <DropdownMenuSubTrigger>Callout</DropdownMenuSubTrigger>
                <DropdownMenuSubContent>
                  {CALLOUT_VARIANTS.map(v => (
                    <DropdownMenuItem key={v.id} onSelect={() => chain().insertKbCallout(v.id).run()}>{v.label}</DropdownMenuItem>
                  ))}
                </DropdownMenuSubContent>
              </DropdownMenuSub>
              <DropdownMenuItem onSelect={() => chain().insertKbDetails().run()}>Details panel<span className="ml-auto text-xs text-muted-foreground">Date / Time / Venue</span></DropdownMenuItem>
              <DropdownMenuItem onSelect={() => chain().insertKbDivider().run()}>Divider</DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
        )}
        <FieldMenu editor={editor} placeholders={placeholders} />
      </Group>
      <Sep />
      <ToolButton
        label="Clear formatting"
        onClick={() => {
          const c = chain();
          ['bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'textStyle', 'kbSpan'].forEach(m => c.unsetMark(m, { extendEmptyMarkRange: true }));
          c.unsetTextAlign().clearBlockStyle().run();
        }}
      >
        <RemoveFormatting />
      </ToolButton>
    </div>
  );
}

function FieldMenu({ editor, placeholders }: { editor: Editor; placeholders: string[] }) {
  const [custom, setCustom] = useState('');
  const insert = (name: string) => editor.chain().focus().insertField(name).insertContent(' ').run();
  return (
    <DropdownMenu>
      <Tip label="Insert recipient field">
        <DropdownMenuTrigger asChild>
          <button type="button" className={tbBtn} aria-label="Insert recipient field" onMouseDown={e => e.preventDefault()}><Braces /><span className="hidden text-xs xl:inline">Field</span><ChevronDown className="!size-3 opacity-60" aria-hidden /></button>
        </DropdownMenuTrigger>
      </Tip>
      <DropdownMenuContent align="end" className="w-60" onCloseAutoFocus={e => e.preventDefault()}>
        <DropdownMenuLabel>Recipient fields</DropdownMenuLabel>
        {placeholders.length === 0 && <p className="px-2 pb-1 text-xs text-muted-foreground">Add recipients to list their columns here, or type a field name below.</p>}
        {placeholders.map(p => (
          <DropdownMenuItem key={p} onSelect={() => insert(p)}><span className="kb-field-preview">{`{${p}}`}</span></DropdownMenuItem>
        ))}
        <DropdownMenuSeparator />
        <form className="flex gap-1 p-1" onSubmit={e => { e.preventDefault(); if (custom.trim()) { insert(custom.trim()); setCustom(''); } }} onKeyDown={e => e.stopPropagation()}>
          <Input value={custom} onChange={e => setCustom(e.target.value.replace(/[{}<>]/g, ''))} placeholder="Field name" aria-label="Custom field name" className="h-8 text-xs" />
          <Button type="submit" size="sm" variant="outline" disabled={!custom.trim()}>Add</Button>
        </form>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}
