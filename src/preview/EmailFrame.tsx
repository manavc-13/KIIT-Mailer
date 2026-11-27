// A sandboxed iframe that renders email HTML at a real viewport width (so media
// queries fire) and sizes itself to its content. Double-buffered: a new document is
// loaded into a hidden frame and swapped in once laid out, so typing in the editor
// never flashes a blank preview or jumps the page height.
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { cn } from '@/lib/utils';

interface EmailFrameProps {
  html: string;
  /** Viewport width in CSS px (the width the email's media queries see). */
  width: number;
  title: string;
  minHeight?: number;
  className?: string;
  background?: string;
}

type Slot = 0 | 1;

function measure(el: HTMLIFrameElement, minHeight: number): number | null {
  let doc: Document | null = null;
  try { doc = el.contentDocument; } catch { return null; }
  if (!doc?.documentElement) return null;
  // Collapse the viewport for the reading so `height:100%` / `100vh` layouts report their
  // content height rather than the current frame height. Restored before any paint.
  const prev = el.style.height;
  el.style.height = '1px';
  const h = Math.max(doc.documentElement.scrollHeight, doc.body?.scrollHeight ?? 0);
  el.style.height = prev;
  return Math.max(minHeight, Math.ceil(h));
}

export function EmailFrame({ html, width, title, minHeight = 120, className, background = '#ffffff' }: EmailFrameProps) {
  const refs = [useRef<HTMLIFrameElement>(null), useRef<HTMLIFrameElement>(null)] as const;
  const [front, setFront] = useState<Slot>(0);
  const [heights, setHeights] = useState<[number, number]>([minHeight, minHeight]);
  const frontRef = useRef<Slot>(0);
  const loaded = useRef<[string | null, string | null]>([null, null]);
  const token = useRef(0);
  const watchers = useRef<[(() => void) | null, (() => void) | null]>([null, null]);
  const minRef = useRef(minHeight);
  minRef.current = minHeight;

  const setHeight = (slot: Slot, h: number | null) => {
    if (h == null) return;
    setHeights(prev => (Math.abs(prev[slot] - h) < 1 ? prev : (slot === 0 ? [h, prev[1]] : [prev[0], h])));
  };

  /** Keep re-measuring the live frame: late images, web fonts, and any reflow. */
  const watch = (slot: Slot) => {
    watchers.current[slot]?.();
    const el = refs[slot].current;
    if (!el) return;
    const remeasure = () => setHeight(slot, measure(el, minRef.current));
    const timers = [150, 500, 1200, 3000].map(ms => window.setTimeout(remeasure, ms));
    let ro: ResizeObserver | null = null;
    let cancelled = false;
    try {
      const doc = el.contentDocument;
      if (doc?.body && typeof ResizeObserver !== 'undefined') {
        ro = new ResizeObserver(() => remeasure());
        ro.observe(doc.body);
      }
      doc?.fonts?.ready.then(() => { if (!cancelled) remeasure(); }).catch(() => {});
      doc?.querySelectorAll('img').forEach(img => {
        if (!img.complete) {
          img.addEventListener('load', remeasure, { once: true });
          img.addEventListener('error', remeasure, { once: true });
        }
      });
    } catch { /* not same-origin: timers still run */ }
    watchers.current[slot] = () => {
      cancelled = true;
      timers.forEach(clearTimeout);
      ro?.disconnect();
    };
  };

  useLayoutEffect(() => {
    const f = frontRef.current;
    const target: Slot = loaded.current[f] == null ? f : ((1 - f) as Slot);
    const el = refs[target].current;
    if (!el) return;
    const t = ++token.current;
    const onLoad = () => {
      if (t !== token.current) return;
      setHeight(target, measure(el, minRef.current));
      if (target !== frontRef.current) {
        watchers.current[frontRef.current]?.();
        frontRef.current = target;
        setFront(target);
      }
      watch(target);
    };
    el.addEventListener('load', onLoad);
    loaded.current[target] = html;
    el.srcdoc = html;
    return () => el.removeEventListener('load', onLoad);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [html]);

  useEffect(() => () => { watchers.current.forEach(w => w?.()); }, []);

  return (
    <div className={cn('relative overflow-hidden', className)} style={{ width, background }}>
      {([0, 1] as const).map(slot => {
        const isFront = slot === front;
        return (
          <iframe
            key={slot}
            ref={refs[slot]}
            title={isFront ? title : `${title} (loading)`}
            sandbox="allow-same-origin"
            referrerPolicy="no-referrer"
            aria-hidden={isFront ? undefined : true}
            tabIndex={isFront ? 0 : -1}
            className="block border-0"
            style={{
              width,
              height: heights[slot],
              background,
              ...(isFront ? {} : { position: 'absolute', top: 0, left: 0, visibility: 'hidden', pointerEvents: 'none' }),
            }}
          />
        );
      })}
    </div>
  );
}
