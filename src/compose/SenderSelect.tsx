// "From" selector: only the sender accounts allotted to the user, with today's usage meter.
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Skeleton } from '@/components/ui/skeleton';
import { cn } from '@/lib/utils';
import type { SenderAccount } from '../../shared/types';
import { useComposeStore } from './store';
import { useSelectedSender } from './model';
import { Meter } from './ui';

export function usageTone(used: number, limit: number): 'ok' | 'warn' | 'danger' {
  if (limit <= 0) return 'danger';
  const r = used / limit;
  return r >= 1 ? 'danger' : r >= 0.8 ? 'warn' : 'ok';
}

export function SenderUsage({ sender, className }: { sender: SenderAccount; className?: string }) {
  const pct = sender.dailyLimit ? (sender.usedToday / sender.dailyLimit) * 100 : 100;
  return (
    <div className={cn('w-40', className)}>
      <div className="mb-1 text-right text-[11.5px] tabular-nums text-muted-foreground">
        <span className="font-medium text-foreground">{sender.usedToday.toLocaleString()}</span> / {sender.dailyLimit.toLocaleString()} today
      </div>
      <Meter percent={pct} tone={usageTone(sender.usedToday, sender.dailyLimit)} label={`Daily sending quota used for ${sender.email}`} />
    </div>
  );
}

export function SenderSelect() {
  const { sender, senders, loading } = useSelectedSender();
  const selectSender = useComposeStore(s => s.selectSender);

  if (loading) return <Skeleton className="h-11 w-72" />;
  if (!senders.length) return null;

  return (
    <div className="flex flex-wrap items-end gap-4">
      <div className="space-y-1">
        <label htmlFor="from-select" className="block text-[11px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">From</label>
        <Select value={sender?.id ?? ''} onValueChange={id => { const s = senders.find(x => x.id === id); if (s) selectSender(s); }}>
          <SelectTrigger id="from-select" className="h-auto min-h-10 w-[19rem] max-w-full bg-card py-1.5 text-left">
            <SelectValue placeholder="Choose a sender account" />
          </SelectTrigger>
          <SelectContent>
            {senders.map(s => (
              <SelectItem key={s.id} value={s.id} className="py-2">
                <span className="flex min-w-0 flex-col">
                  <span className="truncate font-medium">{s.displayName}</span>
                  <span className="truncate text-xs text-muted-foreground">{s.email}</span>
                </span>
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
      {sender && <SenderUsage sender={sender} />}
    </div>
  );
}
