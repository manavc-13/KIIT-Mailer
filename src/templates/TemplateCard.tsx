// One template in the library grid: live thumbnail, details and actions.
import { Copy, MoreHorizontal, PencilLine, Send, Share2, Trash2, Users, Globe2, Lock } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import type { TemplateSummary } from '../../shared/types';
import { relativeTime, themeLabel, visibilityLabel } from './hooks';
import { TemplateThumbnail } from './TemplateThumbnail';

export interface CardActions {
  onUse: (t: TemplateSummary) => void;
  onDuplicate: (t: TemplateSummary) => void;
  onRename: (t: TemplateSummary) => void;
  onShare: (t: TemplateSummary) => void;
  onDelete: (t: TemplateSummary) => void;
}

export function TemplateCard({ t, isAdmin, busy, ...a }: CardActions & { t: TemplateSummary; isAdmin: boolean; busy?: boolean }) {
  const canDelete = t.isOwner || isAdmin;
  const VisIcon = t.visibility === 'shared_all' ? Globe2 : t.visibility === 'shared_users' ? Users : Lock;
  const updated = new Date(t.updatedAt);

  return (
    <article
      className="group flex min-w-0 flex-col overflow-hidden rounded-lg border bg-card shadow-xs transition-[border-color,box-shadow] hover:border-[color-mix(in_oklab,var(--gold)_55%,var(--border))] hover:shadow-sm focus-within:border-[var(--gold)]"
      aria-labelledby={`tpl-${t.id}-name`}
    >
      <button
        type="button"
        onClick={() => a.onUse(t)}
        className="block w-full text-left focus-visible:outline-none"
        tabIndex={-1}
        aria-hidden="true"
      >
        <TemplateThumbnail summary={t} />
      </button>

      <div className="flex flex-1 flex-col gap-2.5 p-4">
        <div className="flex items-start gap-2">
          <div className="min-w-0 flex-1">
            {t.category && (
              <div className="mb-0.5 truncate text-[11px] font-semibold uppercase tracking-[0.12em] text-[var(--gold)]">{t.category}</div>
            )}
            <h3 id={`tpl-${t.id}-name`} className="line-clamp-2 font-serif text-[17px] font-semibold leading-snug text-foreground">
              {t.name}
            </h3>
          </div>
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <Button variant="ghost" size="icon-sm" className="-mr-1.5 -mt-1 shrink-0" aria-label={`More actions for ${t.name}`}>
                <MoreHorizontal />
              </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" className="w-52">
              <DropdownMenuItem onSelect={() => a.onUse(t)}><Send /> Use in Compose</DropdownMenuItem>
              <DropdownMenuItem onSelect={() => a.onDuplicate(t)} disabled={busy}><Copy /> Duplicate to My templates</DropdownMenuItem>
              {t.isOwner && <DropdownMenuItem onSelect={() => a.onRename(t)}><PencilLine /> Rename / category</DropdownMenuItem>}
              {isAdmin && <DropdownMenuItem onSelect={() => a.onShare(t)}><Share2 /> Share…</DropdownMenuItem>}
              {canDelete && (
                <>
                  <DropdownMenuSeparator />
                  <DropdownMenuItem variant="destructive" onSelect={() => a.onDelete(t)}><Trash2 /> Delete…</DropdownMenuItem>
                </>
              )}
            </DropdownMenuContent>
          </DropdownMenu>
        </div>

        <p className="line-clamp-1 text-sm text-muted-foreground" title={t.subject || undefined}>
          <span className="sr-only">Subject: </span>{t.subject || <span className="italic">No subject</span>}
        </p>

        <div className="flex flex-wrap items-center gap-1.5">
          <Badge variant={t.isOwner ? 'secondary' : 'outline'} className="font-normal">
            {t.isOwner ? 'You' : t.ownerName}
          </Badge>
          {t.visibility !== 'private' ? (
            <Badge variant="outline" className="border-[color-mix(in_oklab,var(--gold)_45%,transparent)] bg-[var(--gold-soft)] font-normal text-[var(--accent-foreground)]">
              <VisIcon /> Shared · {visibilityLabel(t.visibility)}
            </Badge>
          ) : t.isOwner ? (
            <Badge variant="outline" className="font-normal text-muted-foreground"><Lock /> Private</Badge>
          ) : null}
          <Badge variant="outline" className="font-normal text-muted-foreground">{themeLabel(t.theme)}</Badge>
          {t.editorMode === 'html' && <Badge variant="outline" className="font-mono text-[10.5px] font-normal text-muted-foreground">HTML</Badge>}
        </div>

        <div className="mt-auto flex items-center justify-between gap-2 border-t pt-3">
          <Tooltip>
            <TooltipTrigger asChild>
              <time dateTime={t.updatedAt} className="text-xs text-muted-foreground" tabIndex={0}>
                Updated {relativeTime(t.updatedAt)}
              </time>
            </TooltipTrigger>
            <TooltipContent>{updated.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}</TooltipContent>
          </Tooltip>
          <div className="flex items-center gap-1">
            {!t.isOwner && (
              <Button variant="ghost" size="sm" onClick={() => a.onDuplicate(t)} disabled={busy} aria-label={`Duplicate ${t.name}`}>
                <Copy /> <span className="hidden sm:inline">Duplicate</span>
              </Button>
            )}
            <Button size="sm" onClick={() => a.onUse(t)} aria-label={`Use ${t.name} in Compose`}>
              <Send /> Use
            </Button>
          </div>
        </div>
      </div>
    </article>
  );
}
