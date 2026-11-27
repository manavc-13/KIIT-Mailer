// Light, realistic client chrome around the email: Gmail web, Gmail app, Apple Mail
// on iPhone, and desktop Outlook. The chrome is decorative (aria-hidden); the
// iframe inside carries the accessible title. Colours here deliberately mimic each
// client and do not follow the app theme.
import type { ReactNode } from 'react';
import {
  Archive, ArrowLeft, BatteryFull, ChevronDown, ChevronLeft, ChevronUp, CornerUpLeft, EllipsisVertical,
  Folder, Forward, Mail, Reply, Search, Signal, SquarePen, Star, Trash2, Wifi, Menu, Printer,
} from 'lucide-react';

export interface ChromeInfo {
  subject: string;
  fromName: string;
  fromEmail: string;
  toLabel: string;
  date: Date;
}

const GMAIL_FONT = "'Google Sans', Roboto, Arial, sans-serif";
const APPLE_FONT = "-apple-system, BlinkMacSystemFont, 'SF Pro Text', 'Helvetica Neue', Helvetica, Arial, sans-serif";
const OUTLOOK_FONT = "'Segoe UI', 'Segoe UI Web (West European)', -apple-system, Arial, sans-serif";

const AVATAR_COLOURS = ['#1a73e8', '#188038', '#a142f4', '#e8710a', '#d93025', '#12b5cb', '#5f6368', '#9334e6'];

export function initialsOf(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return '?';
  return (parts[0][0] + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}

function avatarColour(seed: string): string {
  let h = 0;
  for (let i = 0; i < seed.length; i++) h = (h * 31 + seed.charCodeAt(i)) | 0;
  return AVATAR_COLOURS[Math.abs(h) % AVATAR_COLOURS.length];
}

function Avatar({ name, size, square = false, colour }: { name: string; size: number; square?: boolean; colour?: string }) {
  return (
    <span
      className="inline-grid shrink-0 place-items-center font-medium text-white"
      style={{
        width: size, height: size, borderRadius: square ? 6 : '50%',
        background: colour ?? avatarColour(name), fontSize: Math.round(size * 0.42),
      }}
    >
      {initialsOf(name).slice(0, size < 36 ? 1 : 2)}
    </span>
  );
}

const time = (d: Date) => d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
const shortDate = (d: Date) => d.toLocaleDateString(undefined, { day: 'numeric', month: 'short' });

// ------------------------------------------------------------------ Gmail web (desktop)

export const GMAIL_DESKTOP = { frameWidth: 1120, messageWidth: 920 } as const;

export function GmailDesktopChrome({ info, clipped, children }: { info: ChromeInfo; clipped?: boolean; children: ReactNode }) {
  return (
    <div className="overflow-hidden rounded-xl border border-[#dadce0] bg-[#f6f8fc] text-[#1f1f1f]" style={{ width: GMAIL_DESKTOP.frameWidth, fontFamily: GMAIL_FONT }}>
      <div aria-hidden className="flex h-16 items-center gap-4 px-4">
        <Menu className="size-6 text-[#444746]" strokeWidth={1.75} />
        <span className="w-[150px] text-[22px] text-[#444746]">Mail</span>
        <div className="flex h-12 w-[720px] items-center gap-3 rounded-full bg-[#e9eef6] px-4 text-[16px] text-[#444746]">
          <Search className="size-5" /> Search mail
        </div>
        <span className="ml-auto grid size-8 place-items-center rounded-full bg-[#0b57d0] text-[14px] text-white">M</span>
      </div>
      <div className="flex">
        <div aria-hidden className="flex w-[72px] shrink-0 flex-col items-center gap-5 pt-3 text-[#444746]">
          <span className="grid size-14 place-items-center rounded-2xl bg-[#c2e7ff]"><SquarePen className="size-6" strokeWidth={1.75} /></span>
          <Mail className="size-5" strokeWidth={1.75} />
          <Star className="size-5" strokeWidth={1.75} />
          <Archive className="size-5" strokeWidth={1.75} />
        </div>
        <div className="mb-4 mr-4 min-w-0 flex-1 rounded-2xl bg-white pb-6">
          <div aria-hidden className="flex h-12 items-center gap-6 border-b border-[#f1f3f4] px-4 text-[#444746]">
            <ArrowLeft className="size-5" strokeWidth={1.75} />
            <Archive className="size-5" strokeWidth={1.75} />
            <Trash2 className="size-5" strokeWidth={1.75} />
            <Mail className="size-5" strokeWidth={1.75} />
            <EllipsisVertical className="size-5" strokeWidth={1.75} />
          </div>
          <div className="flex items-center gap-3 pl-[72px] pr-6 pt-5">
            <h3 className="font-normal text-[22px] leading-7 text-[#1f1f1f]" style={{ fontFamily: GMAIL_FONT }}>{info.subject || '(no subject)'}</h3>
            <span className="rounded bg-[#ececec] px-1.5 py-px text-[12px] text-[#444746]">Inbox</span>
            <Printer aria-hidden className="ml-auto size-5 text-[#444746]" strokeWidth={1.75} />
          </div>
          <div className="flex items-start gap-4 px-4 pt-5">
            <Avatar name={info.fromName} size={40} />
            <div className="min-w-0 flex-1 text-[14px] leading-5">
              <div className="flex items-baseline gap-1.5">
                <span className="font-semibold text-[#1f1f1f]">{info.fromName}</span>
                <span className="truncate text-[12px] text-[#5e5e5e]">&lt;{info.fromEmail}&gt;</span>
                <span className="ml-auto flex shrink-0 items-center gap-4 text-[12px] text-[#5e5e5e]">
                  {shortDate(info.date)}, {time(info.date)} (0 minutes ago)
                  <Star aria-hidden className="size-5" strokeWidth={1.5} />
                  <Reply aria-hidden className="size-5" strokeWidth={1.5} />
                  <EllipsisVertical aria-hidden className="size-5" strokeWidth={1.5} />
                </span>
              </div>
              <div className="flex items-center gap-0.5 text-[12px] text-[#5e5e5e]">to {info.toLabel} <ChevronDown aria-hidden className="size-3.5" /></div>
            </div>
          </div>
          <div className="pl-[72px] pr-[40px] pt-4">
            {children}
            {clipped && (
              <div className="mt-3 text-[14px] text-[#1f1f1f]">
                [Message clipped]&nbsp;&nbsp;<span className="text-[#0b57d0] underline">View entire message</span>
              </div>
            )}
          </div>
          <div aria-hidden className="mt-6 flex gap-3 pl-[72px]">
            {[['Reply', Reply], ['Forward', Forward]].map(([label, Icon]) => {
              const I = Icon as typeof Reply;
              return (
                <span key={label as string} className="flex h-10 items-center gap-2 rounded-full border border-[#747775] px-6 text-[14px] text-[#444746]">
                  <I className="size-4" /> {label as string}
                </span>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

// ------------------------------------------------------------------ phones

export const PHONE = { screenWidth: 390, bezel: 10, viewportHeight: 844 } as const;
export const PHONE_FRAME_WIDTH = PHONE.screenWidth + PHONE.bezel * 2;

function StatusBar({ dark }: { dark?: boolean }) {
  return (
    <div aria-hidden className="relative flex h-[47px] items-center justify-between px-8 pt-1 text-[15px] font-semibold" style={{ color: dark ? '#fff' : '#000', fontFamily: APPLE_FONT }}>
      <span>9:41</span>
      <span className="absolute left-1/2 top-[11px] h-[30px] w-[118px] -translate-x-1/2 rounded-full bg-black" />
      <span className="flex items-center gap-1.5">
        <Signal className="size-4" strokeWidth={2.5} /><Wifi className="size-4" strokeWidth={2.5} /><BatteryFull className="size-5" strokeWidth={2} />
      </span>
    </div>
  );
}

/** A 390 px iPhone-sized screen in a dark bezel. Grows with content; marks the first screenful. */
export function PhoneShell({ dark, children, screenBg }: { dark?: boolean; children: ReactNode; screenBg?: string }) {
  return (
    <div className="rounded-[52px] bg-[#1d1d1f] shadow-[0_1px_3px_rgba(0,0,0,0.25)]" style={{ width: PHONE_FRAME_WIDTH, padding: PHONE.bezel }}>
      <div className="relative overflow-hidden rounded-[42px]" style={{ width: PHONE.screenWidth, minHeight: 720, background: screenBg ?? (dark ? '#1f1f1f' : '#fff') }}>
        <StatusBar dark={dark} />
        {children}
        <div aria-hidden className="pointer-events-none absolute inset-x-0 flex items-center gap-2 px-2" style={{ top: PHONE.viewportHeight }}>
          <span className="h-px flex-1 border-t border-dashed border-[#c5973a]" />
          <span className="rounded-sm bg-[#c5973a] px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-white">End of first screen</span>
        </div>
        <div aria-hidden className="flex justify-center pb-2 pt-4"><span className="h-[5px] w-[134px] rounded-full" style={{ background: dark ? '#fff' : '#000' }} /></div>
      </div>
    </div>
  );
}

export function GmailMobileChrome({ info, clipped, dark, children }: { info: ChromeInfo; clipped?: boolean; dark?: boolean; children: ReactNode }) {
  const fg = dark ? '#e3e3e3' : '#1f1f1f';
  const muted = dark ? '#a8a8a8' : '#5e5e5e';
  return (
    <PhoneShell dark={dark}>
      <div style={{ fontFamily: GMAIL_FONT, color: fg }}>
        <div aria-hidden className="flex h-14 items-center gap-6 px-4" style={{ color: muted }}>
          <ArrowLeft className="size-6" strokeWidth={1.75} />
          <span className="ml-auto flex gap-6">
            <Archive className="size-5" strokeWidth={1.75} /><Trash2 className="size-5" strokeWidth={1.75} />
            <Mail className="size-5" strokeWidth={1.75} /><EllipsisVertical className="size-5" strokeWidth={1.75} />
          </span>
        </div>
        <div className="flex items-start gap-2 px-4 pb-3">
          <h3 className="flex-1 text-[22px] font-normal leading-7" style={{ fontFamily: GMAIL_FONT, color: fg }}>
            {info.subject || '(no subject)'}{' '}
            <span className="relative -top-0.5 rounded px-1.5 py-px align-middle text-[12px]" style={{ background: dark ? '#3c3c3c' : '#ececec', color: muted }}>Inbox</span>
          </h3>
          <Star aria-hidden className="mt-1 size-5 shrink-0" style={{ color: muted }} strokeWidth={1.5} />
        </div>
        <div className="flex items-center gap-3 px-4 pb-3">
          <Avatar name={info.fromName} size={40} />
          <div className="min-w-0 flex-1 text-[14px] leading-5">
            <div className="flex items-baseline gap-2">
              <span className="truncate font-semibold">{info.fromName}</span>
              <span className="shrink-0 text-[12px]" style={{ color: muted }}>{time(info.date)}</span>
            </div>
            <div className="flex items-center gap-0.5 text-[12px]" style={{ color: muted }}>to {info.toLabel} <ChevronDown aria-hidden className="size-3.5" /></div>
          </div>
          <CornerUpLeft aria-hidden className="size-5" style={{ color: muted }} strokeWidth={1.75} />
          <EllipsisVertical aria-hidden className="size-5" style={{ color: muted }} strokeWidth={1.75} />
        </div>
        {children}
        {clipped && (
          <div className="px-4 pt-3 text-[14px]">[Message clipped]&nbsp;&nbsp;<span className="text-[#0b57d0] underline">View entire message</span></div>
        )}
        <div aria-hidden className="flex gap-3 px-4 pb-2 pt-5">
          {['Reply', 'Forward'].map(l => (
            <span key={l} className="flex h-10 flex-1 items-center justify-center rounded-full border text-[14px]" style={{ borderColor: dark ? '#5f5f5f' : '#747775', color: muted }}>{l}</span>
          ))}
        </div>
      </div>
    </PhoneShell>
  );
}

export function AppleMailChrome({ info, children }: { info: ChromeInfo; children: ReactNode }) {
  const blue = '#007aff';
  return (
    <PhoneShell>
      <div style={{ fontFamily: APPLE_FONT, color: '#000' }}>
        <div aria-hidden className="flex h-11 items-center px-2 text-[17px]" style={{ color: blue }}>
          <ChevronLeft className="size-7" strokeWidth={2.25} /> Inbox
          <span className="ml-auto flex gap-5 pr-3"><ChevronUp className="size-6" strokeWidth={2} /><ChevronDown className="size-6" strokeWidth={2} /></span>
        </div>
        <div className="flex items-start gap-3 px-4 pt-2">
          <Avatar name={info.fromName} size={40} colour="#a3a8b3" />
          <div className="min-w-0 flex-1 text-[15px] leading-5">
            <div className="flex items-baseline gap-2">
              <span className="truncate font-semibold">{info.fromName}</span>
              <span className="ml-auto shrink-0 text-[14px] text-[#8e8e93]">{time(info.date)}</span>
            </div>
            <div className="truncate text-[14px] text-[#8e8e93]">To: {info.toLabel}</div>
          </div>
        </div>
        <h3 className="px-4 pb-3 pt-3 text-[20px] font-semibold leading-6" style={{ fontFamily: APPLE_FONT }}>{info.subject || '(no subject)'}</h3>
        <div className="mx-4 mb-0 border-t border-[#d1d1d6]" />
        {children}
        <div aria-hidden className="flex items-center justify-between border-t border-[#d1d1d6] px-6 pt-3" style={{ color: blue }}>
          <Trash2 className="size-6" strokeWidth={1.75} /><Folder className="size-6" strokeWidth={1.75} />
          <CornerUpLeft className="size-6" strokeWidth={1.75} /><SquarePen className="size-6" strokeWidth={1.75} />
        </div>
      </div>
    </PhoneShell>
  );
}

// ------------------------------------------------------------------ Outlook for Windows

export const OUTLOOK_DESKTOP = { frameWidth: 1000, messageWidth: 920 } as const;

export function OutlookChrome({ info, children }: { info: ChromeInfo; children: ReactNode }) {
  return (
    <div className="overflow-hidden rounded-lg border border-[#d1d1d1] bg-white text-[#242424]" style={{ width: OUTLOOK_DESKTOP.frameWidth, fontFamily: OUTLOOK_FONT }}>
      <div aria-hidden className="flex h-9 items-center gap-3 bg-[#0f6cbd] px-4 text-[13px] text-white">
        <Menu className="size-4" /> Inbox · {info.toLabel}
        <span className="ml-auto flex h-6 w-[320px] items-center gap-2 rounded bg-white/90 px-2 text-[#616161]"><Search className="size-3.5" /> Search</span>
      </div>
      <div aria-hidden className="flex h-10 items-center gap-5 border-b border-[#e0e0e0] bg-[#f5f5f5] px-4 text-[13px] text-[#424242]">
        <span className="flex items-center gap-1.5"><Reply className="size-4" /> Reply</span>
        <span className="flex items-center gap-1.5"><Reply className="size-4" /> Reply all</span>
        <span className="flex items-center gap-1.5"><Forward className="size-4" /> Forward</span>
        <span className="flex items-center gap-1.5"><Trash2 className="size-4" /> Delete</span>
      </div>
      <div className="px-10 pb-3 pt-5">
        <h3 className="text-[20px] font-semibold leading-7" style={{ fontFamily: OUTLOOK_FONT }}>{info.subject || '(no subject)'}</h3>
        <div className="mt-4 flex items-start gap-3">
          <Avatar name={info.fromName} size={40} colour="#c19c00" />
          <div className="min-w-0 flex-1 text-[14px] leading-5">
            <div className="flex items-baseline gap-1.5">
              <span className="font-semibold">{info.fromName}</span>
              <span className="truncate text-[#616161]">&lt;{info.fromEmail}&gt;</span>
              <span className="ml-auto shrink-0 text-[12px] text-[#616161]">
                {info.date.toLocaleDateString(undefined, { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' })} {time(info.date)}
              </span>
            </div>
            <div className="text-[12px] text-[#616161]">To: {info.toLabel}</div>
          </div>
        </div>
      </div>
      <div className="px-10 pb-8">{children}</div>
    </div>
  );
}
