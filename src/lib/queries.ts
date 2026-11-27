// Shared TanStack Query hooks used by several feature areas.
import { useQuery } from '@tanstack/react-query';
import { api } from './api';
import type { SenderAccount, TemplateSummary, UploadConfig } from '../../shared/types';

/** Sender accounts the signed-in user may send from. */
export function useMySenders() {
  return useQuery({
    queryKey: ['senders', 'mine'],
    queryFn: async () => (await api.get<{ senders: SenderAccount[] }>('/senders')).senders,
  });
}

/** Templates visible to the signed-in user (own + shared). */
export function useTemplates() {
  return useQuery({
    queryKey: ['templates'],
    queryFn: async () => (await api.get<{ templates: TemplateSummary[] }>('/templates')).templates,
  });
}

export function useUploadConfig() {
  return useQuery({
    queryKey: ['uploads', 'config'],
    queryFn: () => api.get<UploadConfig>('/uploads/config'),
    staleTime: Infinity,
  });
}
