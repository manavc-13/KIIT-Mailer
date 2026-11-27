// Template data access for the library page, picker and save dialog.
import { useQuery, useQueryClient, type QueryClient } from '@tanstack/react-query';
import { api } from '@/lib/api';
import { useAuth } from '@/lib/auth';
import type { AdminUser, Template, TemplateSummary, ThemeId } from '../../shared/types';
import { THEMES } from '../../shared/themes';

export { useTemplates } from '@/lib/queries';

/** All templates in the system (admin only; GET /templates?scope=all). */
export function useAllTemplates(enabled: boolean) {
  return useQuery({
    queryKey: ['templates', 'scope-all'],
    queryFn: async () => (await api.get<{ templates: TemplateSummary[] }>('/templates?scope=all')).templates,
    enabled,
  });
}

/**
 * Full template detail. Keyed by updatedAt (outside the ['templates'] prefix) so
 * refreshing the list only refetches templates that actually changed.
 */
export const detailKey = (id: string, updatedAt?: string) => ['template-detail', id, updatedAt ?? ''] as const;

export function useTemplateDetail(summary: Pick<TemplateSummary, 'id' | 'updatedAt'> | null, enabled = true) {
  return useQuery({
    queryKey: detailKey(summary?.id ?? '', summary?.updatedAt),
    queryFn: async () => (await api.get<{ template: Template }>(`/templates/${summary!.id}`)).template,
    enabled: enabled && !!summary,
    staleTime: Infinity,
    gcTime: 10 * 60 * 1000,
  });
}

export function fetchTemplate(qc: QueryClient, summary: Pick<TemplateSummary, 'id' | 'updatedAt'>): Promise<Template> {
  return qc.fetchQuery({
    queryKey: detailKey(summary.id, summary.updatedAt),
    queryFn: async () => (await api.get<{ template: Template }>(`/templates/${summary.id}`)).template,
    staleTime: Infinity,
  });
}

/** Seed the detail cache after a save so the next open is instant. */
export function primeTemplate(qc: QueryClient, t: Template) {
  qc.setQueryData(detailKey(t.id, t.updatedAt), t);
}

export function useInvalidateTemplates() {
  const qc = useQueryClient();
  return () => qc.invalidateQueries({ queryKey: ['templates'] });
}

/** Users for the admin share dialog. */
export function useAdminUsers(enabled: boolean) {
  const { isAdmin } = useAuth();
  return useQuery({
    queryKey: ['admin', 'users'],
    queryFn: async () => (await api.get<{ users: AdminUser[] }>('/admin/users')).users,
    enabled: enabled && isAdmin,
  });
}

const FALLBACK_THEME_LABELS: Record<ThemeId, string> = {
  heritage: 'Heritage', meridian: 'Meridian', letterhead: 'Letterhead', slate: 'Slate', plain: 'Plain',
};

export function themeLabel(id: ThemeId): string {
  return THEMES.find(t => t.id === id)?.label ?? FALLBACK_THEME_LABELS[id] ?? id;
}

export function visibilityLabel(v: TemplateSummary['visibility']): string {
  return v === 'shared_all' ? 'Everyone' : v === 'shared_users' ? 'Specific people' : 'Private';
}

/** Distinct, sorted categories. */
export function categoriesOf(list: TemplateSummary[] | undefined): string[] {
  const set = new Map<string, string>();
  for (const t of list ?? []) {
    const c = t.category?.trim();
    if (c && !set.has(c.toLowerCase())) set.set(c.toLowerCase(), c);
  }
  return [...set.values()].sort((a, b) => a.localeCompare(b));
}

const rtf = new Intl.RelativeTimeFormat('en', { numeric: 'auto' });
export function relativeTime(isoDate: string, now = Date.now()): string {
  const diff = (new Date(isoDate).getTime() - now) / 1000;
  const abs = Math.abs(diff);
  if (abs < 60) return 'just now';
  if (abs < 3600) return rtf.format(Math.round(diff / 60), 'minute');
  if (abs < 86400) return rtf.format(Math.round(diff / 3600), 'hour');
  if (abs < 86400 * 7) return rtf.format(Math.round(diff / 86400), 'day');
  return new Date(isoDate).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
}
