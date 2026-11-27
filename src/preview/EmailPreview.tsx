// <EmailPreview>: renders the server's real send output per client and device.
import { useId, useMemo, useRef, useState, type KeyboardEvent, type ReactNode } from 'react';
import {
  ChevronLeft, ChevronRight, Columns2, FileText, Loader2, MailX, Monitor, Moon, RotateCcw, Smartphone, TriangleAlert,
} from 'lucide-react';
import { cn } from '@/lib/utils';
import { Button } from '@/components/ui/button';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import type { RenderRequest, RenderResult } from '../../shared/types';
import { htmlForClient, type PreviewClientId } from './transforms';
import { EmailFrame } from './EmailFrame';
import { ScaledFrame } from './ScaledFrame';
import {
  AppleMailChrome, GMAIL_DESKTOP, GmailDesktopChrome, GmailMobileChrome, OUTLOOK_DESKTOP, OutlookChrome,
  PHONE, PHONE_FRAME_WIDTH, type ChromeInfo,
} from './chrome';
import { ChecksPanel } from './ChecksPanel';
import { useRenderPreview } from './useRenderPreview';
import { readPref, writePref } from './prefs';

export interface EmailPreviewProps {
  /** null = nothing to preview yet (e.g. no sender chosen). */
  request: RenderRequest | null;
  /** Recipient navigation (shown when the list has > 1 row). */
  recipientIndex?: number;
  recipientCount?: number;
  onRecipientIndexChange?: (i: number) => void;
  /** Sender display for the client chrome (avatar, "from" line). */
  fromName?: string;
  fromEmail?: string;
  toLabel?: string;
  /**
   * Receives every successful render (the review step uses lint + size) with the key of the request
   * it was rendered from (renderRequestKey), so callers can ignore a result for an older request.
   */
  onRendered?: (result: RenderResult, requestKey: string) => void;
  /** 'rail' = compact docked panel in Compose; 'full' = large view in Review. */
  layout?: 'rail' | 'full';
  className?: string;
  /** Message shown when `request` is null. Default: "Choose a sender to preview". */
  emptyMessage?: string;
  /** Hide the Checks panel (e.g. when the host shows lint elsewhere). */
  hideChecks?: boolean;
  /** Open the Checks panel by default (until the user toggles it). */
  checksDefaultOpen?: boolean;
}

interface ClientDef {
  id: PreviewClientId;
  label: string;
  icon: typeof Monitor;
  badge?: string;
  note: ReactNode;
  tone?: 'warn';
}

const CLIENTS: ClientDef[] = [
  {
    id: 'gmail-desktop', label: 'Gmail · Desktop', icon: Monitor, badge: 'G',
    note: `Gmail on the web at its real reading width (${GMAIL_DESKTOP.messageWidth} px). Web fonts show as their fallbacks, as Gmail does.`,
  },
  {
    id: 'gmail-mobile', label: 'Gmail · Mobile', icon: Smartphone, badge: 'G',
    note: `Gmail app on a ${PHONE.screenWidth} px phone screen. Mobile styles apply; web fonts show as their fallbacks.`,
  },
  {
    id: 'apple-iphone', label: 'Apple Mail · iPhone', icon: Smartphone, badge: 'A',
    note: `Apple Mail on a ${PHONE.screenWidth} px iPhone screen. Full fidelity: web fonts load and mobile styles apply.`,
  },
  {
    id: 'outlook-windows', label: 'Outlook · Windows', icon: Monitor, badge: 'O', tone: 'warn',
    note: 'Approximation. Desktop Outlook uses Microsoft Word to render email; send a test to check exact rendering.',
  },
  {
    id: 'dark', label: 'Dark mode (reference)', icon: Moon, tone: 'warn',
    note: 'Reference only. The email opts out of dark mode re-theming, but some apps ignore that and invert colours like this.',
  },
  {
    id: 'text', label: 'Plain text', icon: FileText,
    note: 'The plain-text part sent with every message. Some clients, watches and screen readers show this instead of the HTML.',
  },
];

const CLIENT_IDS = CLIENTS.map(c => c.id);
const isClientId = (v: string): v is PreviewClientId => (CLIENT_IDS as string[]).includes(v);

function ClientGlyph({ def, className }: { def: ClientDef; className?: string }) {
  const Icon = def.icon;
  return (
    <span className={cn('relative inline-flex', className)} aria-hidden>
      <Icon className="size-4" strokeWidth={1.75} />
      {def.badge && (
        <span className="absolute -bottom-1 -right-1.5 grid size-3 place-items-center rounded-[3px] bg-current text-[8px] font-bold leading-none">
          <span className="text-background">{def.badge}</span>
        </span>
      )}
    </span>
  );
}

function ClientTabs({ value, onChange, compact, panelId }: {
  value: PreviewClientId; onChange: (c: PreviewClientId) => void; compact: boolean; panelId: string;
}) {
  const refs = useRef<(HTMLButtonElement | null)[]>([]);
  const onKey = (e: KeyboardEvent, i: number) => {
    let next = -1;
    if (e.key === 'ArrowRight') next = (i + 1) % CLIENTS.length;
    else if (e.key === 'ArrowLeft') next = (i - 1 + CLIENTS.length) % CLIENTS.length;
    else if (e.key === 'Home') next = 0;
    else if (e.key === 'End') next = CLIENTS.length - 1;
    if (next < 0) return;
    e.preventDefault();
    onChange(CLIENTS[next].id);
    refs.current[next]?.focus();
  };
  return (
    <div role="tablist" aria-label="Email client" className="inline-flex flex-wrap items-center rounded-md border bg-muted p-0.5">
      {CLIENTS.map((c, i) => {
        const selected = c.id === value;
        const btn = (
          <button
            ref={el => { refs.current[i] = el; }}
            type="button"
            role="tab"
            aria-selected={selected}
            aria-controls={panelId}
            aria-label={compact ? c.label : undefined}
            tabIndex={selected ? 0 : -1}
            onClick={() => onChange(c.id)}
            onKeyDown={e => onKey(e, i)}
            className={cn(
              'inline-flex h-8 items-center gap-2 rounded-[5px] text-[12.5px] font-medium text-muted-foreground transition-colors',
              'hover:text-foreground focus-visible:outline-2 focus-visible:outline-offset-1 focus-visible:outline-ring',
              compact ? 'w-9 justify-center' : 'px-2.5',
              selected && 'bg-card text-foreground shadow-[0_1px_2px_rgba(12,26,40,0.12)]',
            )}
          >
            <ClientGlyph def={c} className={!compact && c.badge ? 'mr-1' : undefined} />
            {!compact && <span className="whitespace-nowrap">{c.label.replace(' (reference)', '')}</span>}
          </button>
        );
        return compact ? (
          <Tooltip key={c.id}>
            <TooltipTrigger asChild>{btn}</TooltipTrigger>
            <TooltipContent side="bottom">{c.label}</TooltipContent>
          </Tooltip>
        ) : <span key={c.id} className="contents">{btn}</span>;
      })}
    </div>
  );
}

function RecipientNav({ index, count, onChange }: { index: number; count: number; onChange?: (i: number) => void }) {
  if (count <= 1) return null;
  const go = (d: number) => onChange?.((index + d + count) % count);
  return (
    <div className="inline-flex items-center gap-0.5 rounded-md border bg-card" role="group" aria-label="Previewed recipient">
      <Button variant="ghost" size="icon-sm" className="size-8" onClick={() => go(-1)} aria-label="Previous recipient" disabled={!onChange}>
        <ChevronLeft />
      </Button>
      <span className="min-w-[4.5rem] text-center text-[12.5px] tabular-nums" aria-live="polite">
        {index + 1} / {count}
      </span>
      <Button variant="ghost" size="icon-sm" className="size-8" onClick={() => go(1)} aria-label="Next recipient" disabled={!onChange}>
        <ChevronRight />
      </Button>
    </div>
  );
}

function Stage({ children, className }: { children: ReactNode; className?: string }) {
  return <div className={cn('rounded-lg border bg-[#e9e7e1] p-3 dark:bg-[#0a1017] sm:p-4', className)}>{children}</div>;
}

function PlainTextPane({ text }: { text: string }) {
  return (
    <pre
      tabIndex={0}
      aria-label="Plain-text version"
      className="mx-auto max-w-[78ch] whitespace-pre-wrap break-words rounded-md border bg-card px-5 py-4 font-mono text-[13px] leading-[1.6] text-card-foreground"
    >
      {text || '(empty)'}
    </pre>
  );
}

function ScaleNote({ scale }: { scale: number }) {
  if (scale >= 0.995) return null;
  return <span className="text-[11.5px] tabular-nums text-muted-foreground">Shown at {Math.round(scale * 100)}%</span>;
}

/** One client view: transformed HTML inside its chrome, scaled to fit. */
function ClientView({ client, result, info, layout, maxPhoneWidth }: {
  client: PreviewClientId; result: RenderResult; info: ChromeInfo; layout: 'rail' | 'full'; maxPhoneWidth?: number;
}) {
  const [scale, setScale] = useState(1);
  const prepared = useMemo(
    () => (client === 'text' ? null : htmlForClient(client, result.html, result.sizeBytes)),
    [client, result.html, result.sizeBytes],
  );
  const title = `Email preview: ${CLIENTS.find(c => c.id === client)?.label}`;

  if (client === 'text' || !prepared) return <PlainTextPane text={result.text} />;

  const phoneCap = maxPhoneWidth ?? (layout === 'full' ? PHONE_FRAME_WIDTH : undefined);
  let body: ReactNode;
  let natural: number;
  let cap: number | undefined;
  switch (client) {
    case 'gmail-desktop':
      natural = GMAIL_DESKTOP.frameWidth;
      body = (
        <GmailDesktopChrome info={info} clipped={prepared.clipped}>
          <EmailFrame html={prepared.html} width={GMAIL_DESKTOP.messageWidth} title={title} />
        </GmailDesktopChrome>
      );
      break;
    case 'outlook-windows':
      natural = OUTLOOK_DESKTOP.frameWidth;
      body = (
        <OutlookChrome info={info}>
          <EmailFrame html={prepared.html} width={OUTLOOK_DESKTOP.messageWidth} title={title} />
        </OutlookChrome>
      );
      break;
    case 'apple-iphone':
      natural = PHONE_FRAME_WIDTH; cap = phoneCap;
      body = (
        <AppleMailChrome info={info}>
          <EmailFrame html={prepared.html} width={PHONE.screenWidth} title={title} />
        </AppleMailChrome>
      );
      break;
    case 'dark':
      natural = PHONE_FRAME_WIDTH; cap = phoneCap;
      body = (
        <GmailMobileChrome info={info} dark>
          <EmailFrame html={prepared.html} width={PHONE.screenWidth} title={title} background="#1f1f1f" />
        </GmailMobileChrome>
      );
      break;
    case 'gmail-mobile':
    default:
      natural = PHONE_FRAME_WIDTH; cap = phoneCap;
      body = (
        <GmailMobileChrome info={info} clipped={prepared.clipped}>
          <EmailFrame html={prepared.html} width={PHONE.screenWidth} title={title} />
        </GmailMobileChrome>
      );
  }
  return (
    <div>
      <ScaledFrame naturalWidth={natural} maxRenderedWidth={cap} onScale={setScale}>{body}</ScaledFrame>
      <div className="mt-1.5 flex justify-end"><ScaleNote scale={scale} /></div>
    </div>
  );
}

function comparePair(client: PreviewClientId): [PreviewClientId, PreviewClientId] | null {
  if (client === 'dark' || client === 'text') return null;
  return [client === 'outlook-windows' ? 'outlook-windows' : 'gmail-desktop', client === 'apple-iphone' ? 'apple-iphone' : 'gmail-mobile'];
}

export function EmailPreview(props: EmailPreviewProps) {
  const {
    request, recipientIndex = 0, recipientCount = 0, onRecipientIndexChange, fromName, fromEmail, toLabel,
    onRendered, layout = 'rail', className, emptyMessage, hideChecks, checksDefaultOpen,
  } = props;
  const panelId = useId();
  const state = useRenderPreview(request, { onRendered });
  // The tab choice is remembered per layout. The narrow Compose rail defaults to the phone view,
  // which fits at close to 100%; a 920 px desktop render squeezed into the rail is too small to read.
  const defaultClient: PreviewClientId = layout === 'rail' ? 'gmail-mobile' : 'gmail-desktop';
  const clientPrefKey = `client.${layout}`;
  const [client, setClientState] = useState<PreviewClientId>(() => {
    const v = readPref(clientPrefKey, defaultClient);
    return isClientId(v) ? v : defaultClient;
  });
  const [compare, setCompareState] = useState(() => readPref('compare', '0') === '1');
  const setClient = (c: PreviewClientId) => { setClientState(c); writePref(clientPrefKey, c); };
  const setCompare = (v: boolean) => { setCompareState(v); writePref('compare', v ? '1' : '0'); };

  const result = state.result;
  // One timestamp per rendered message so the chrome's clock doesn't tick on every re-render.
  const date = useMemo(() => new Date(), [result]);
  const senderName = fromName?.trim() || (fromEmail ? fromEmail.split('@')[0] : 'KIIT IQAC');
  const info: ChromeInfo = {
    subject: result?.subject ?? request?.subject ?? '',
    fromName: senderName,
    fromEmail: fromEmail || 'sender@kiit.ac.in',
    toLabel: toLabel || 'me',
    date,
  };
  const def = CLIENTS.find(c => c.id === client)!;
  const compact = layout === 'rail';
  const pair = layout === 'full' && compare ? comparePair(client) : null;

  let content: ReactNode;
  if (state.problem?.kind === 'no-sender') {
    content = (
      <Stage className="grid min-h-[320px] place-items-center text-center">
        <div className="max-w-xs">
          <MailX className="mx-auto size-8 text-muted-foreground" strokeWidth={1.5} aria-hidden />
          <p className="mt-3 font-serif text-[17px] font-semibold">{request ? state.problem.message : (emptyMessage ?? state.problem.message)}</p>
          <p className="mt-1 text-[13px] text-muted-foreground">The preview is rendered with the chosen account's signature and letterhead, exactly as it will be sent.</p>
        </div>
      </Stage>
    );
  } else if (!result) {
    content = state.problem ? (
      <Stage className="grid min-h-[320px] place-items-center text-center">
        <div role="alert" className="max-w-sm">
          <TriangleAlert className="mx-auto size-7 text-warning" strokeWidth={1.75} aria-hidden />
          <p className="mt-3 text-[14px] font-medium">The preview could not be rendered.</p>
          <p className="mt-1 text-[13px] text-muted-foreground">{state.problem.message}</p>
          <Button variant="outline" size="sm" className="mt-4" onClick={state.refetch}><RotateCcw /> Try again</Button>
        </div>
      </Stage>
    ) : (
      <Stage className="grid min-h-[320px] place-items-center">
        <span className="flex items-center gap-2 text-[13px] text-muted-foreground"><Loader2 className="size-4 animate-spin" aria-hidden /> Rendering preview…</span>
      </Stage>
    );
  } else if (pair) {
    content = (
      <Stage>
        <div className="grid items-start gap-4 xl:grid-cols-[minmax(0,1fr)_minmax(0,340px)]">
          <ClientView client={pair[0]} result={result} info={info} layout={layout} />
          <ClientView client={pair[1]} result={result} info={info} layout={layout} maxPhoneWidth={340} />
        </div>
      </Stage>
    );
  } else {
    content = (
      <Stage className={client === 'text' ? 'bg-muted dark:bg-muted' : undefined}>
        <ClientView client={client} result={result} info={info} layout={layout} />
      </Stage>
    );
  }

  return (
    <div className={cn('flex min-w-0 flex-col gap-3', className)}>
      <div className="flex flex-wrap items-center gap-2">
        <ClientTabs value={client} onChange={setClient} compact={compact} panelId={panelId} />
        {layout === 'full' && (
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="outline"
                size="sm"
                aria-pressed={compare}
                disabled={!comparePair(client)}
                onClick={() => setCompare(!compare)}
                className={cn(compare && 'border-gold bg-gold-soft text-accent-foreground')}
              >
                <Columns2 /> Desktop + mobile
              </Button>
            </TooltipTrigger>
            <TooltipContent side="bottom">Show the desktop and phone views side by side</TooltipContent>
          </Tooltip>
        )}
        <div className="ml-auto flex items-center gap-2">
          <span aria-live="polite" className="text-[12px] text-muted-foreground">
            {state.updating && result && (
              <span className="inline-flex items-center gap-1.5"><Loader2 className="size-3.5 animate-spin" aria-hidden /> Updating…</span>
            )}
          </span>
          <RecipientNav index={recipientIndex} count={recipientCount} onChange={onRecipientIndexChange} />
        </div>
      </div>

      <p className={cn('flex items-start gap-1.5 text-[12.5px] leading-5', def.tone === 'warn' ? 'text-warning' : 'text-muted-foreground')}>
        {def.tone === 'warn' && <TriangleAlert className="mt-0.5 size-3.5 shrink-0" aria-hidden />}
        <span>{pair ? 'Desktop and phone side by side. ' : ''}{def.note}</span>
      </p>

      {result && state.problem?.kind === 'error' && (
        <div role="alert" className="flex items-center gap-2 rounded-md border border-warning/40 bg-warning-soft px-3 py-2 text-[12.5px] text-warning">
          <TriangleAlert className="size-4 shrink-0" aria-hidden />
          <span className="min-w-0 flex-1">Showing the last good preview. {state.problem.message}</span>
          <Button variant="ghost" size="xs" onClick={state.refetch}><RotateCcw /> Retry</Button>
        </div>
      )}

      <div id={panelId} role="tabpanel" aria-label={def.label} className={cn('transition-opacity', state.stale && state.updating && 'opacity-90')}>
        {content}
      </div>

      {!hideChecks && result && state.problem?.kind !== 'no-sender' && (
        <ChecksPanel result={result} defaultOpen={checksDefaultOpen ?? layout === 'full'} />
      )}
    </div>
  );
}
