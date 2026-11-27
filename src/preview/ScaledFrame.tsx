// Renders children at a fixed natural width and scales them down (CSS transform)
// to fit the available width, with a wrapper whose height matches the scaled
// content. The email inside keeps its real viewport width, so media queries are
// evaluated exactly as on the device.
import { useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { cn } from '@/lib/utils';

interface ScaledFrameProps {
  naturalWidth: number;
  /** Upper bound for the scale (1 = never enlarge). */
  maxScale?: number;
  /** Optional cap on the rendered (scaled) width, e.g. to keep phones compact. */
  maxRenderedWidth?: number;
  children: ReactNode;
  className?: string;
  /** Reports the scale in use, e.g. for a "62%" label. */
  onScale?: (scale: number) => void;
}

export function ScaledFrame({ naturalWidth, maxScale = 1, maxRenderedWidth, children, className, onScale }: ScaledFrameProps) {
  const outerRef = useRef<HTMLDivElement>(null);
  const innerRef = useRef<HTMLDivElement>(null);
  const [outerWidth, setOuterWidth] = useState(0);
  const [innerHeight, setInnerHeight] = useState(0);

  useLayoutEffect(() => {
    const outer = outerRef.current;
    const inner = innerRef.current;
    if (!outer || !inner) return;
    const update = () => {
      setOuterWidth(outer.clientWidth);
      setInnerHeight(inner.offsetHeight);
    };
    update();
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(update) : null;
    ro?.observe(outer);
    ro?.observe(inner);
    // Auto-sized iframes change their style attribute; catch that even when no frame is painted yet.
    const mo = new MutationObserver(update);
    mo.observe(inner, { attributes: true, attributeFilter: ['style'], subtree: true, childList: true });
    window.addEventListener('resize', update);
    return () => { ro?.disconnect(); mo.disconnect(); window.removeEventListener('resize', update); };
  }, []);

  const available = maxRenderedWidth ? Math.min(outerWidth, maxRenderedWidth) : outerWidth;
  const scale = outerWidth > 0 ? Math.min(maxScale, available / naturalWidth) : 1;
  const scaledWidth = naturalWidth * scale;
  const offsetX = Math.max(0, (outerWidth - scaledWidth) / 2);

  const lastScale = useRef(0);
  useLayoutEffect(() => {
    if (onScale && Math.abs(lastScale.current - scale) > 0.001) {
      lastScale.current = scale;
      onScale(scale);
    }
  }, [scale, onScale]);

  return (
    <div ref={outerRef} className={cn('relative w-full overflow-hidden', className)} style={{ height: innerHeight * scale || undefined }}>
      <div
        ref={innerRef}
        className="absolute left-0 top-0 origin-top-left"
        style={{ width: naturalWidth, transform: `translateX(${offsetX}px) scale(${scale})` }}
      >
        {children}
      </div>
    </div>
  );
}
