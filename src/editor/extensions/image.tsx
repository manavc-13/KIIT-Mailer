// Email image node: hosted src, alt, px width (<= 600), alignment, and a transient
// `rehost` flag for pasted images that are still being uploaded. React node view with
// width presets, a drag-resize handle, alt text and an "uploading" overlay.
import { useState } from 'react';
import { NodeViewWrapper, ReactNodeViewRenderer, type ReactNodeViewProps } from '@tiptap/react';
import Image from '@tiptap/extension-image';
import { AlignCenter, AlignLeft, AlignRight, Loader2, Trash2, Type } from 'lucide-react';
import { MAX_IMAGE_WIDTH } from '../constants';
import { lengthToPx, parseStyle } from '../style-utils';
import { AltTextDialog } from '../dialogs';

const clampWidth = (w: number) => Math.max(24, Math.min(MAX_IMAGE_WIDTH, Math.round(w)));

function ImageView({ node, updateAttributes, deleteNode, selected, editor }: ReactNodeViewProps) {
  const { src, alt, width, align, rehost } = node.attrs as { src: string; alt: string; width: number | null; align: string | null; rehost: boolean };
  const [altOpen, setAltOpen] = useState(false);
  const [dragW, setDragW] = useState<number | null>(null);
  const editable = editor.isEditable;
  const shown = dragW ?? width;

  const onLoad = (e: React.SyntheticEvent<HTMLImageElement>) => {
    if (!width && !rehost && editable) {
      const nat = e.currentTarget.naturalWidth;
      if (nat) updateAttributes({ width: clampWidth(nat) });
    }
  };

  const startResize = (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const img = (e.currentTarget.parentElement?.querySelector('img') as HTMLImageElement | null);
    const startX = e.clientX;
    const startW = img?.getBoundingClientRect().width || width || 300;
    const dir = align === 'right' ? -1 : align === 'center' ? 2 : 1;
    let latest = startW;
    const move = (ev: PointerEvent) => { latest = clampWidth(startW + (ev.clientX - startX) * dir); setDragW(latest); };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      setDragW(null);
      updateAttributes({ width: latest });
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const justify = align === 'center' ? 'center' : align === 'right' ? 'flex-end' : 'flex-start';
  const btn = 'inline-flex h-7 min-w-7 items-center justify-center rounded px-1.5 text-xs font-medium hover:bg-white/15 focus-visible:outline focus-visible:outline-2 focus-visible:outline-[var(--gold)]';

  return (
    <NodeViewWrapper className="kb-image-node" data-align={align || 'left'} style={{ display: 'flex', justifyContent: justify }}>
      <div className={`kb-image-frame${selected ? ' is-selected' : ''}`} style={{ width: shown ? `${shown}px` : undefined, maxWidth: '100%' }}>
        <img src={src} alt={alt} draggable={false} onLoad={onLoad} style={{ display: 'block', width: '100%', height: 'auto', opacity: rehost ? 0.45 : 1 }} />
        {rehost && (
          <span className="kb-image-badge" role="status"><Loader2 className="size-3.5 animate-spin" aria-hidden /> Uploading…</span>
        )}
        {!alt && !rehost && editable && (
          <button type="button" className="kb-image-alt-hint" onClick={() => setAltOpen(true)}>Add alt text</button>
        )}
        {selected && editable && !rehost && (
          <>
            <div className="kb-image-toolbar" contentEditable={false} role="toolbar" aria-label="Image options">
              {[25, 50, 75, 100].map(p => (
                <button key={p} type="button" className={btn} aria-label={`Width ${p}%`} title={`${p}% of the email width`} onClick={() => updateAttributes({ width: clampWidth((MAX_IMAGE_WIDTH * p) / 100) })}>{p}%</button>
              ))}
              <span className="mx-0.5 h-4 w-px bg-white/25" aria-hidden />
              <button type="button" className={btn} aria-label="Align left" title="Align left" onClick={() => updateAttributes({ align: null })}><AlignLeft className="size-3.5" /></button>
              <button type="button" className={btn} aria-label="Align centre" title="Align centre" onClick={() => updateAttributes({ align: 'center' })}><AlignCenter className="size-3.5" /></button>
              <button type="button" className={btn} aria-label="Align right" title="Align right" onClick={() => updateAttributes({ align: 'right' })}><AlignRight className="size-3.5" /></button>
              <span className="mx-0.5 h-4 w-px bg-white/25" aria-hidden />
              <button type="button" className={btn} aria-label="Edit alt text" title="Alt text" onClick={() => setAltOpen(true)}><Type className="size-3.5" /></button>
              <button type="button" className={btn} aria-label="Remove image" title="Remove image" onClick={deleteNode}><Trash2 className="size-3.5" /></button>
            </div>
            <span className="kb-image-handle" onPointerDown={startResize} aria-hidden title="Drag to resize" />
            <span className="kb-image-size" aria-hidden>{shown ? `${Math.round(shown)}px` : 'auto'}</span>
          </>
        )}
      </div>
      <AltTextDialog open={altOpen} onOpenChange={setAltOpen} initial={alt || ''} onSubmit={v => updateAttributes({ alt: v })} />
    </NodeViewWrapper>
  );
}

export const EmailImage = Image.extend({
  draggable: true,

  addAttributes() {
    return {
      src: { default: null },
      alt: { default: '' },
      title: { default: null },
      width: {
        default: null,
        parseHTML: el => {
          const a = el.getAttribute('width');
          const w = a && /^\d+(\.\d+)?$/.test(a) ? parseFloat(a) : lengthToPx(parseStyle(el.getAttribute('style')).width);
          return w && w > 0 ? clampWidth(w) : null;
        },
        renderHTML: attrs => (attrs.width ? { width: attrs.width } : {}),
      },
      align: {
        default: null,
        parseHTML: el => {
          const a = el.getAttribute('data-align');
          return a === 'center' || a === 'right' ? a : null;
        },
        renderHTML: attrs => (attrs.align ? { 'data-align': attrs.align } : {}),
      },
      rehost: {
        default: false,
        parseHTML: el => el.getAttribute('data-kb-rehost') === '1',
        renderHTML: attrs => (attrs.rehost ? { 'data-kb-rehost': '1' } : {}),
      },
    };
  },

  addNodeView() {
    return ReactNodeViewRenderer(ImageView);
  },
}).configure({ inline: false, allowBase64: true, resize: false });
