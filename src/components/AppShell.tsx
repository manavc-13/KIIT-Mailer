// App chrome: fixed navy sidebar (navigation + account menu) and the page area.
import { NavLink, Outlet, useNavigate } from 'react-router-dom';
import {
  PenLine, LibraryBig, History, Users, KeyRound, ScrollText, ShieldBan, LogOut, Moon, Sun, Lock, ChevronsUpDown, Menu,
} from 'lucide-react';
import { useQueryClient } from '@tanstack/react-query';
import { signedOut, useAuth } from '@/lib/auth';
import { api } from '@/lib/api';
import { useTheme } from '@/lib/theme';
import { cn } from '@/lib/utils';
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuLabel, DropdownMenuSeparator, DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import { ActiveJobBanner } from '@/components/ActiveJobBanner';

const MAIN_NAV = [
  { to: '/compose', label: 'Compose', icon: PenLine },
  { to: '/templates', label: 'Templates', icon: LibraryBig },
  { to: '/history', label: 'Sent history', icon: History },
];
const ADMIN_NAV = [
  { to: '/admin/users', label: 'Users', icon: Users },
  { to: '/admin/senders', label: 'Sender accounts', icon: KeyRound },
  { to: '/admin/audit', label: 'Audit log', icon: ScrollText },
  { to: '/admin/suppressions', label: 'Suppression list', icon: ShieldBan },
];

function NavItem({ to, label, icon: Icon }: { to: string; label: string; icon: typeof PenLine }) {
  return (
    <NavLink
      to={to}
      className={({ isActive }) => cn(
        'group flex items-center gap-3 rounded-md px-3 py-2 text-[13.5px] font-medium transition-colors',
        isActive
          ? 'bg-sidebar-accent text-white shadow-[inset_2px_0_0_var(--gold)]'
          : 'text-sidebar-foreground/80 hover:bg-sidebar-accent/60 hover:text-white',
      )}
    >
      <Icon className="size-4 opacity-80" strokeWidth={1.75} />
      {label}
    </NavLink>
  );
}

export function Monogram({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 32 32" className={className} aria-hidden="true">
      <rect width="32" height="32" rx="6" fill="#0c1a28" />
      <path d="M7 10.5h18v11H7z" fill="none" stroke="#C5973A" strokeWidth="1.8" />
      <path d="M7.5 11l8.5 6.5 8.5-6.5" fill="none" stroke="#C5973A" strokeWidth="1.8" />
    </svg>
  );
}

export function AppShell() {
  const { user, isAdmin } = useAuth();
  const { theme, toggle } = useTheme();
  const qc = useQueryClient();
  const navigate = useNavigate();

  // Used by both the sidebar account menu and the mobile menu.
  async function signOut() {
    try { await api.post('/auth/logout'); } catch { /* already signed out */ }
    // Set ME_KEY to null (notifies AuthProvider, unlike qc.clear()) and drop every other cached query.
    signedOut(qc);
    navigate('/login', { replace: true });
  }

  return (
    <div className="flex h-full min-h-0">
      <aside className="hidden w-60 shrink-0 flex-col border-r border-sidebar-border bg-sidebar text-sidebar-foreground md:flex">
        <div className="flex items-center gap-2.5 px-5 pb-6 pt-5">
          <Monogram className="size-8 rounded-md ring-1 ring-white/10" />
          <div className="leading-tight">
            <div className="font-serif text-[17px] font-semibold text-white">KIIT Mailer</div>
            <div className="text-[11px] uppercase tracking-[0.14em] text-sidebar-muted">IQAC · Official mail</div>
          </div>
        </div>
        <nav className="flex flex-1 flex-col gap-0.5 px-3">
          {MAIN_NAV.map(n => <NavItem key={n.to} {...n} />)}
          {isAdmin && (
            <>
              <div className="mb-1 mt-6 px-3 text-[10.5px] font-semibold uppercase tracking-[0.16em] text-sidebar-muted">Administration</div>
              {ADMIN_NAV.map(n => <NavItem key={n.to} {...n} />)}
            </>
          )}
        </nav>
        <div className="border-t border-sidebar-border p-3">
          <DropdownMenu>
            <DropdownMenuTrigger className="flex w-full items-center gap-3 rounded-md px-2 py-2 text-left hover:bg-sidebar-accent focus-visible:outline-2 focus-visible:outline-[var(--gold)]">
              <span className="grid size-8 place-items-center rounded-full bg-[var(--gold)] text-[13px] font-semibold text-[#0c1a28]">
                {(user?.name || '?').charAt(0).toUpperCase()}
              </span>
              <span className="min-w-0 flex-1">
                <span className="block truncate text-[13px] font-medium text-white">{user?.name}</span>
                <span className="block truncate text-[11.5px] text-sidebar-muted">{user?.email}</span>
              </span>
              <ChevronsUpDown className="size-4 text-sidebar-muted" />
            </DropdownMenuTrigger>
            <DropdownMenuContent side="top" align="start" className="w-56">
              <DropdownMenuLabel className="text-xs font-normal text-muted-foreground">
                Signed in as {user?.role === 'admin' ? 'Administrator' : 'Sender'}
              </DropdownMenuLabel>
              <DropdownMenuSeparator />
              <DropdownMenuItem onSelect={() => navigate('/change-password')}><Lock /> Change password</DropdownMenuItem>
              <DropdownMenuItem onSelect={toggle}>{theme === 'dark' ? <Sun /> : <Moon />} {theme === 'dark' ? 'Light' : 'Dark'} appearance</DropdownMenuItem>
              <DropdownMenuSeparator />
              <DropdownMenuItem onSelect={signOut}><LogOut /> Sign out</DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
        </div>
      </aside>

      <div className="flex min-w-0 flex-1 flex-col">
        {/* Compact top bar below the md breakpoint */}
        <header className="flex items-center gap-3 border-b bg-sidebar px-4 py-3 text-white md:hidden">
          <Monogram className="size-7 rounded" />
          <span className="font-serif font-semibold">KIIT Mailer</span>
          <DropdownMenu>
            <DropdownMenuTrigger
              aria-label="Open menu"
              className="ml-auto grid size-9 place-items-center rounded-md text-white/90 hover:bg-sidebar-accent focus-visible:outline-2 focus-visible:outline-[var(--gold)]"
            >
              <Menu className="size-5" />
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" className="w-60">
              <DropdownMenuLabel className="truncate text-xs font-normal text-muted-foreground">{user?.email}</DropdownMenuLabel>
              <DropdownMenuSeparator />
              {MAIN_NAV.map(({ to, label, icon: Icon }) => (
                <DropdownMenuItem key={to} onSelect={() => navigate(to)}><Icon /> {label}</DropdownMenuItem>
              ))}
              {isAdmin && (
                <>
                  <DropdownMenuSeparator />
                  <DropdownMenuLabel className="text-[10.5px] uppercase tracking-[0.14em] text-muted-foreground">Administration</DropdownMenuLabel>
                  {ADMIN_NAV.map(({ to, label, icon: Icon }) => (
                    <DropdownMenuItem key={to} onSelect={() => navigate(to)}><Icon /> {label}</DropdownMenuItem>
                  ))}
                </>
              )}
              <DropdownMenuSeparator />
              <DropdownMenuItem onSelect={() => navigate('/change-password')}><Lock /> Change password</DropdownMenuItem>
              <DropdownMenuItem onSelect={toggle}>{theme === 'dark' ? <Sun /> : <Moon />} {theme === 'dark' ? 'Light' : 'Dark'} appearance</DropdownMenuItem>
              <DropdownMenuItem onSelect={signOut}><LogOut /> Sign out</DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
        </header>
        <ActiveJobBanner />
        <main className="min-h-0 flex-1 overflow-y-auto">
          <Outlet />
        </main>
      </div>
    </div>
  );
}

/** Standard page header used by every page: serif title, optional description and actions. */
export function PageHeader({ title, description, actions, className }: {
  title: string; description?: React.ReactNode; actions?: React.ReactNode; className?: string;
}) {
  return (
    <div className={cn('flex flex-wrap items-end justify-between gap-4 border-b bg-card/60 px-6 pb-5 pt-7 lg:px-10', className)}>
      <div className="min-w-0">
        <h1 className="text-[26px] font-semibold leading-tight text-foreground">{title}</h1>
        {description && <p className="mt-1 max-w-2xl text-sm text-muted-foreground">{description}</p>}
      </div>
      {actions && <div className="flex flex-wrap items-center gap-2">{actions}</div>}
    </div>
  );
}

/** Standard padded page body. */
export function PageBody({ children, className }: { children: React.ReactNode; className?: string }) {
  return <div className={cn('px-6 py-6 lg:px-10', className)}>{children}</div>;
}
