// Auth context: who is signed in. Backed by GET /api/auth/me via TanStack Query.
import { createContext, useContext, useEffect, type ReactNode } from 'react';
import { useQuery, useQueryClient, type QueryClient } from '@tanstack/react-query';
import { api, ApiClientError, setUnauthorizedHandler } from './api';
import type { MeResponse, PublicUser } from '../../shared/types';

interface AuthState {
  user: PublicUser | null;
  loading: boolean;
  isAdmin: boolean;
  /** Re-fetch /auth/me (call after login, logout, password change). */
  refresh: () => Promise<void>;
  /** Set the user directly from a login response without a round trip. */
  setUser: (u: PublicUser | null) => void;
}

const AuthContext = createContext<AuthState | null>(null);
export const ME_KEY = ['auth', 'me'] as const;

/** Drop every cached query except the auth ones (templates, history, senders, admin lists…). */
export function clearUserQueries(qc: QueryClient) {
  qc.removeQueries({ predicate: q => q.queryKey[0] !== 'auth' });
}

/**
 * Mark the session as signed out. Setting ME_KEY (rather than qc.clear()) notifies the
 * AuthProvider observer, so route guards see the signed-out state immediately.
 */
export function signedOut(qc: QueryClient) {
  qc.setQueryData(ME_KEY, null);
  clearUserQueries(qc);
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const qc = useQueryClient();
  const { data, isLoading } = useQuery({
    queryKey: ME_KEY,
    queryFn: async () => {
      try { return (await api.get<MeResponse>('/auth/me')).user; }
      catch (e) { if (e instanceof ApiClientError && e.status === 401) return null; throw e; }
    },
    staleTime: 5 * 60 * 1000,
    retry: false,
  });

  useEffect(() => {
    setUnauthorizedHandler(() => signedOut(qc));
  }, [qc]);

  const value: AuthState = {
    user: data ?? null,
    loading: isLoading,
    isAdmin: data?.role === 'admin',
    refresh: async () => { await qc.invalidateQueries({ queryKey: ME_KEY }); },
    setUser: (u) => {
      const prev = qc.getQueryData<PublicUser | null>(ME_KEY);
      // A different person on the same tab must never see the previous user's cached data.
      if ((prev?.id ?? null) !== (u?.id ?? null)) clearUserQueries(qc);
      qc.setQueryData(ME_KEY, u);
    },
  };
  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function useAuth(): AuthState {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used inside <AuthProvider>');
  return ctx;
}
