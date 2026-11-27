// TanStack Query hooks for the admin screens.
import { keepPreviousData, useQuery } from '@tanstack/react-query';
import { api } from '@/lib/api';
import type { AdminUser, SenderAccountAdmin } from '../../shared/types';

export const ADMIN_USERS_KEY = ['admin', 'users'] as const;
export const ADMIN_SENDERS_KEY = ['admin', 'senders'] as const;
export const SUPPRESSIONS_KEY = ['admin', 'suppressions'] as const;
/** The server returns at most this many suppressions per request (newest first). */
export const SUPPRESSIONS_LIST_MAX = 5000;

export interface Suppression { email: string; reason: 'unsubscribed' | 'bounced' | 'manual'; note: string | null; at: string }

export function useAdminUsers() {
  return useQuery({
    queryKey: ADMIN_USERS_KEY,
    queryFn: async () => (await api.get<{ users: AdminUser[] }>('/admin/users')).users,
  });
}

export function useAdminSenders() {
  return useQuery({
    queryKey: ADMIN_SENDERS_KEY,
    queryFn: async () => (await api.get<{ senders: SenderAccountAdmin[] }>('/admin/senders')).senders,
  });
}

/** Suppressions, newest first. `q` searches addresses on the server (needed beyond the 5,000 cap). */
export function useSuppressions(q = '', opts: { enabled?: boolean } = {}) {
  const term = q.trim().toLowerCase();
  return useQuery({
    queryKey: [...SUPPRESSIONS_KEY, term],
    queryFn: async ({ signal }) => (await api.get<{ suppressions: Suppression[] }>(
      term ? `/admin/suppressions?q=${encodeURIComponent(term)}` : '/admin/suppressions', { signal },
    )).suppressions,
    enabled: opts.enabled ?? true,
    placeholderData: term ? keepPreviousData : undefined,
  });
}
