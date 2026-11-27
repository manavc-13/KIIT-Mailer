// "Sending in progress" banner: polls GET /api/jobs/active (on mount and every 30 s) while signed
// in, and links back to the live progress page. Hidden on /send/* and when nothing is active.
import { Link, useLocation } from 'react-router-dom';
import { useQuery } from '@tanstack/react-query';
import { Loader2, PauseCircle, ArrowRight } from 'lucide-react';
import { api } from '@/lib/api';
import { useAuth } from '@/lib/auth';
import type { Job } from '../../shared/types';

export function ActiveJobBanner() {
  const { user } = useAuth();
  const { pathname } = useLocation();
  const { data } = useQuery({
    queryKey: ['jobs', 'active'],
    queryFn: async () => (await api.get<{ jobs: Job[] }>('/jobs/active')).jobs,
    enabled: !!user && !user.mustChangePassword,
    refetchInterval: 30_000,
    refetchOnMount: 'always',
    retry: false,
  });

  const jobs = data ?? [];
  if (!user || !jobs.length || pathname.startsWith('/send/')) return null;

  // Prefer a job that is actually sending over a paused one.
  const job = jobs.find(j => j.status === 'sending') ?? jobs.find(j => j.status === 'queued') ?? jobs[0];
  const processed = job.sentCount + job.failedCount + job.skippedCount;
  const paused = job.status === 'paused';
  const others = jobs.length - 1;

  return (
    <div role="status" className="flex flex-wrap items-center gap-x-3 gap-y-1 border-b border-gold/30 bg-gold-soft px-6 py-2 text-[13px] text-accent-foreground lg:px-10">
      {paused ? <PauseCircle className="size-4 shrink-0 text-gold" /> : <Loader2 className="size-4 shrink-0 animate-spin text-gold" />}
      <span className="min-w-0 flex-1 truncate">
        <span className="font-medium">{paused ? 'Sending paused' : 'Sending in progress'}:</span>{' '}
        “{job.subject}” ({processed.toLocaleString()}/{job.totalMessages.toLocaleString()})
        {others > 0 && <span className="text-muted-foreground"> and {others} more</span>}
      </span>
      <Link
        to={`/send/${job.id}`}
        className="inline-flex items-center gap-1 rounded px-1.5 py-0.5 font-semibold underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-[var(--gold)]"
      >
        {paused ? 'Resume' : 'Open'} <ArrowRight className="size-3.5" />
      </Link>
    </div>
  );
}
