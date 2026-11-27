// Live, scaled-down preview of a template in a fully sandboxed iframe.
// The full template is fetched only once the card scrolls into view.
import { useEffect, useMemo, useRef, useState } from 'react';
import { FileWarning } from 'lucide-react';
import type { Template, TemplateSummary } from '../../shared/types';
import { renderShell } from '../../shared/themes';
import { Skeleton } from '@/components/ui/skeleton';
import { cn } from '@/lib/utils';
import { useTemplateDetail } from './hooks';

/** Width the email is laid out at before scaling (typical email body width plus margins). */
const FRAME_WIDTH = 680;
const FRAME_HEIGHT = 900;

export function templatePreviewHtml(t: Template): string {
  if (t.editorMode === 'html') return t.html;
  return renderShell({ theme: t.theme, bodyHtml: t.html, signatureHtml: null, subject: t.subject });
}

function useInView<T extends Element>(rootMargin = '200px') {
  const ref = useRef<T>(null);
  const [inView, setInView] = useState(false);
  useEffect(() => {
    const el = ref.current;
    if (!el || inView) return;
    if (typeof IntersectionObserver === 'undefined') { setInView(true); return; }
    const io = new IntersectionObserver((entries) => {
      if (entries.some(e => e.isIntersecting)) { setInView(true); io.disconnect(); }
    }, { rootMargin });
    io.observe(el);
    return () => io.disconnect();
  }, [inView, rootMargin]);
  return { ref, inView };
}

function useWidth<T extends Element>(ref: React.RefObject<T | null>) {
  const [w, setW] = useState(0);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    setW(el.clientWidth);
    if (typeof ResizeObserver === 'undefined') return;
    const ro = new ResizeObserver(([e]) => setW(e.contentRect.width));
    ro.observe(el);
    return () => ro.disconnect();
  }, [ref]);
  return w;
}

export function TemplateThumbnail({ summary, className }: { summary: TemplateSummary; className?: string }) {
  const { ref, inView } = useInView<HTMLDivElement>();
  const width = useWidth(ref);
  const { data, isError } = useTemplateDetail(summary, inView);
  const srcDoc = useMemo(() => (data ? templatePreviewHtml(data) : ''), [data]);
  const scale = width > 0 ? width / FRAME_WIDTH : 0.35;

  return (
    <div
      ref={ref}
      className={cn('relative aspect-[4/3] overflow-hidden border-b bg-[#f4f3ef]', className)}
      aria-hidden="true"
    >
      {isError ? (
        <div className="grid h-full place-items-center text-xs text-muted-foreground">
          <span className="flex items-center gap-1.5"><FileWarning className="size-4" /> Preview unavailable</span>
        </div>
      ) : !data ? (
        <div className="flex h-full flex-col gap-2 p-5">
          <Skeleton className="h-4 w-1/3" />
          <Skeleton className="h-3 w-5/6" />
          <Skeleton className="h-3 w-4/6" />
          <Skeleton className="h-3 w-3/4" />
        </div>
      ) : !srcDoc.trim() ? (
        <div className="grid h-full place-items-center text-xs text-muted-foreground">Empty template</div>
      ) : (
        <iframe
          title={`Preview of ${summary.name}`}
          srcDoc={srcDoc}
          sandbox=""
          loading="lazy"
          tabIndex={-1}
          className="pointer-events-none absolute left-0 top-0 origin-top-left border-0 bg-white"
          style={{ width: FRAME_WIDTH, height: FRAME_HEIGHT, transform: `scale(${scale})` }}
        />
      )}
      {/* Soft fade so a cut-off email reads as a thumbnail, not a broken frame. */}
      <div className="pointer-events-none absolute inset-x-0 bottom-0 h-10 bg-gradient-to-t from-[#f4f3ef] to-transparent" />
    </div>
  );
}
