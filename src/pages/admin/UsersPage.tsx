// Admin: users. Create accounts (one-time temporary password), edit name/role,
// allot sender accounts, reset passwords and disable/enable access.
import { useMemo, useState, type FormEvent } from 'react';
import { useMutation, useQueryClient } from '@tanstack/react-query';
import {
  KeyRound, MoreHorizontal, Pencil, RotateCcw, ShieldCheck, UserCheck, UserPlus, UserX, Users,
} from 'lucide-react';
import { toast } from 'sonner';
import { PageHeader, PageBody } from '@/components/AppShell';
import { api, errorMessage } from '@/lib/api';
import { ME_KEY, useAuth } from '@/lib/auth';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Checkbox } from '@/components/ui/checkbox';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import {
  Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle,
} from '@/components/ui/dialog';
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import { ConfirmDialog, EmptyState, Pill, SearchBox, TableSkeleton, TempPasswordReveal } from '@/admin/components';
import { ADMIN_SENDERS_KEY, ADMIN_USERS_KEY, useAdminSenders, useAdminUsers } from '@/admin/queries';
import { formatDateTime, relativeTime } from '@/admin/format';
import type { AdminUser, CreateUserRequest, Role, SenderAccountAdmin, UpdateUserRequest } from '../../../shared/types';

// ------------------------------------------------------------------ pieces

function RoleBadge({ role }: { role: Role }) {
  return role === 'admin'
    ? <Pill tone="gold"><ShieldCheck /> Administrator</Pill>
    : <Pill>Sender</Pill>;
}

function StatusBadge({ user }: { user: AdminUser }) {
  if (user.disabled) return <Pill tone="danger">Disabled</Pill>;
  if (user.mustChangePassword) {
    return <Pill tone="warning" title="Has a temporary password that has not been changed yet">Awaiting first sign-in</Pill>;
  }
  return <Pill tone="success">Active</Pill>;
}

function SenderChecklist({ senders, value, onChange, idPrefix }: {
  senders: SenderAccountAdmin[]; value: string[]; onChange: (ids: string[]) => void; idPrefix: string;
}) {
  if (!senders.length) {
    return <p className="rounded-md border border-dashed px-3 py-4 text-sm text-muted-foreground">No sender accounts have been added yet. Add them on the Sender accounts page.</p>;
  }
  const toggle = (id: string, on: boolean) => onChange(on ? [...value, id] : value.filter(v => v !== id));
  return (
    <div className="max-h-64 divide-y overflow-y-auto rounded-md border bg-card">
      {senders.map(s => {
        const id = `${idPrefix}-${s.id}`;
        return (
          <label key={s.id} htmlFor={id} className="flex cursor-pointer items-start gap-3 px-3 py-2.5 hover:bg-muted/60">
            <Checkbox id={id} className="mt-0.5" checked={value.includes(s.id)} onCheckedChange={c => toggle(s.id, c === true)} />
            <span className="min-w-0 flex-1">
              <span className="block truncate text-sm font-medium">{s.displayName}</span>
              <span className="block truncate text-xs text-muted-foreground">{s.email}</span>
            </span>
            {s.disabled && <Pill tone="danger">Disabled</Pill>}
          </label>
        );
      })}
    </div>
  );
}

// ------------------------------------------------------------------ dialogs

function CreateUserDialog({ open, onOpenChange, senders }: {
  open: boolean; onOpenChange: (o: boolean) => void; senders: SenderAccountAdmin[];
}) {
  const qc = useQueryClient();
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [role, setRole] = useState<Role>('sender');
  const [senderIds, setSenderIds] = useState<string[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [created, setCreated] = useState<{ user: AdminUser; tempPassword: string } | null>(null);

  const reset = () => { setName(''); setEmail(''); setRole('sender'); setSenderIds([]); setError(null); setCreated(null); };
  /** Every way out (Cancel, Done, Esc, outside click, ×) closes and clears the form for next time. */
  const close = () => {
    if (create.isPending) return;
    onOpenChange(false);
    setTimeout(reset, 200); // after the close animation
  };

  const create = useMutation({
    mutationFn: (req: CreateUserRequest) => api.post<{ user: AdminUser; tempPassword: string }>('/admin/users', req),
    onSuccess: (res) => {
      setCreated(res);
      qc.invalidateQueries({ queryKey: ADMIN_USERS_KEY });
      qc.invalidateQueries({ queryKey: ADMIN_SENDERS_KEY });
    },
    onError: (err) => setError(errorMessage(err)),
  });

  function submit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    if (!name.trim()) return setError('Enter the person’s name.');
    if (!/^\S+@\S+\.\S+$/.test(email.trim())) return setError('Enter a valid email address.');
    create.mutate({ name: name.trim(), email: email.trim(), role, senderAccountIds: senderIds });
  }

  return (
    <Dialog open={open} onOpenChange={(o) => { if (!o) close(); else onOpenChange(true); }}>
      <DialogContent className="sm:max-w-xl">
        {created ? (
          <>
            <DialogHeader>
              <DialogTitle className="font-serif text-xl">User created</DialogTitle>
              <DialogDescription>
                {created.user.name} can now sign in at this site with {created.user.email} and the temporary password below.
              </DialogDescription>
            </DialogHeader>
            <TempPasswordReveal password={created.tempPassword} who={created.user.name} />
            <DialogFooter>
              <Button onClick={close}>Done</Button>
            </DialogFooter>
          </>
        ) : (
          <form onSubmit={submit} noValidate className="grid gap-5">
            <DialogHeader>
              <DialogTitle className="font-serif text-xl">Add a user</DialogTitle>
              <DialogDescription>
                A temporary password is generated for the new user. They choose their own the first time they sign in.
              </DialogDescription>
            </DialogHeader>
            {error && <p role="alert" className="rounded-md border border-destructive/30 bg-destructive/5 px-3 py-2 text-sm text-destructive">{error}</p>}
            <div className="grid gap-4 sm:grid-cols-2">
              <div className="space-y-2">
                <Label htmlFor="nu-name">Full name</Label>
                <Input id="nu-name" value={name} onChange={e => setName(e.target.value)} autoFocus maxLength={120} />
              </div>
              <div className="space-y-2">
                <Label htmlFor="nu-email">Email address</Label>
                <Input id="nu-email" type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="name@kiit.ac.in" />
              </div>
            </div>
            <fieldset className="space-y-2">
              <legend className="mb-2 text-sm font-medium">Role</legend>
              <RadioGroup value={role} onValueChange={v => setRole(v as Role)} className="grid gap-2 sm:grid-cols-2">
                {([
                  ['sender', 'Sender', 'Composes and sends from the accounts allotted below.'],
                  ['admin', 'Administrator', 'Manages users, sender accounts and the audit log. Can send from every account.'],
                ] as const).map(([v, label, help]) => (
                  <label key={v} htmlFor={`nu-role-${v}`} className="flex cursor-pointer gap-3 rounded-md border bg-card p-3 has-[[data-state=checked]]:border-[var(--gold)] has-[[data-state=checked]]:bg-gold-soft">
                    <RadioGroupItem id={`nu-role-${v}`} value={v} className="mt-0.5" />
                    <span>
                      <span className="block text-sm font-medium">{label}</span>
                      <span className="block text-xs text-muted-foreground">{help}</span>
                    </span>
                  </label>
                ))}
              </RadioGroup>
            </fieldset>
            <div className="space-y-2">
              <div className="text-sm font-medium" id="nu-senders-label">Sender accounts this user may send from</div>
              {role === 'admin' && <p className="text-xs text-muted-foreground">Administrators can use every enabled account; allotments matter only if the role changes later.</p>}
              <div aria-labelledby="nu-senders-label">
                <SenderChecklist senders={senders} value={senderIds} onChange={setSenderIds} idPrefix="nu-s" />
              </div>
            </div>
            <DialogFooter>
              <Button type="button" variant="ghost" onClick={close} disabled={create.isPending}>Cancel</Button>
              <Button type="submit" disabled={create.isPending}>{create.isPending ? 'Creating…' : 'Create user'}</Button>
            </DialogFooter>
          </form>
        )}
      </DialogContent>
    </Dialog>
  );
}

function EditUserDialog({ user, onClose, isSelf }: { user: AdminUser | null; onClose: () => void; isSelf: boolean }) {
  const qc = useQueryClient();
  const [name, setName] = useState(user?.name ?? '');
  const [role, setRole] = useState<Role>(user?.role ?? 'sender');
  const [error, setError] = useState<string | null>(null);

  const save = useMutation({
    mutationFn: (req: UpdateUserRequest) => api.patch<{ user: AdminUser }>(`/admin/users/${user!.id}`, req),
    onSuccess: (res) => {
      qc.invalidateQueries({ queryKey: ADMIN_USERS_KEY });
      // Renaming yourself: refresh the signed-in user so the sidebar shows the new name.
      if (isSelf) qc.invalidateQueries({ queryKey: ME_KEY });
      toast.success(`Saved changes to ${res.user.name}.`);
      onClose();
    },
    onError: (err) => setError(errorMessage(err)),
  });

  if (!user) return null;
  return (
    <Dialog open onOpenChange={(o) => { if (!o && !save.isPending) onClose(); }}>
      <DialogContent>
        <form
          noValidate className="grid gap-5"
          onSubmit={(e) => {
            e.preventDefault();
            if (!name.trim()) return setError('Enter a name.');
            const req: UpdateUserRequest = {};
            if (name.trim() !== user.name) req.name = name.trim();
            if (role !== user.role) req.role = role;
            if (!Object.keys(req).length) return onClose();
            save.mutate(req);
          }}
        >
          <DialogHeader>
            <DialogTitle className="font-serif text-xl">Edit {user.name}</DialogTitle>
            <DialogDescription>{user.email}</DialogDescription>
          </DialogHeader>
          {error && <p role="alert" className="rounded-md border border-destructive/30 bg-destructive/5 px-3 py-2 text-sm text-destructive">{error}</p>}
          <div className="space-y-2">
            <Label htmlFor="eu-name">Full name</Label>
            <Input id="eu-name" value={name} onChange={e => setName(e.target.value)} maxLength={120} />
          </div>
          <fieldset className="space-y-2">
            <legend className="mb-2 text-sm font-medium">Role</legend>
            <RadioGroup value={role} onValueChange={v => setRole(v as Role)} className="flex gap-6" disabled={isSelf}>
              <label htmlFor="eu-role-sender" className="flex items-center gap-2 text-sm"><RadioGroupItem id="eu-role-sender" value="sender" /> Sender</label>
              <label htmlFor="eu-role-admin" className="flex items-center gap-2 text-sm"><RadioGroupItem id="eu-role-admin" value="admin" /> Administrator</label>
            </RadioGroup>
            {isSelf
              ? <p className="text-xs text-muted-foreground">You cannot change your own role.</p>
              : role !== user.role && <p className="text-xs text-muted-foreground">Changing the role signs {user.name} out so the change takes effect immediately.</p>}
          </fieldset>
          <DialogFooter>
            <Button type="button" variant="ghost" onClick={onClose} disabled={save.isPending}>Cancel</Button>
            <Button type="submit" disabled={save.isPending}>{save.isPending ? 'Saving…' : 'Save changes'}</Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

function SenderAccessDialog({ user, senders, onClose }: { user: AdminUser | null; senders: SenderAccountAdmin[]; onClose: () => void }) {
  const qc = useQueryClient();
  const [ids, setIds] = useState<string[]>(user?.senderAccountIds ?? []);
  const save = useMutation({
    mutationFn: () => api.put<{ user: AdminUser }>(`/admin/users/${user!.id}/senders`, { senderAccountIds: ids }),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ADMIN_USERS_KEY });
      qc.invalidateQueries({ queryKey: ADMIN_SENDERS_KEY });
      toast.success(`Updated sender access for ${user!.name}.`);
      onClose();
    },
    onError: (err) => toast.error(errorMessage(err)),
  });
  if (!user) return null;
  return (
    <Dialog open onOpenChange={(o) => { if (!o && !save.isPending) onClose(); }}>
      <DialogContent className="sm:max-w-lg">
        <DialogHeader>
          <DialogTitle className="font-serif text-xl">Sender access for {user.name}</DialogTitle>
          <DialogDescription>
            Tick the accounts {user.name} may send from. Only these appear in their From list.
            {user.role === 'admin' && ' As an administrator they can already use every enabled account.'}
          </DialogDescription>
        </DialogHeader>
        <SenderChecklist senders={senders} value={ids} onChange={setIds} idPrefix="sa" />
        <DialogFooter>
          <Button variant="ghost" onClick={onClose} disabled={save.isPending}>Cancel</Button>
          <Button onClick={() => save.mutate()} disabled={save.isPending}>{save.isPending ? 'Saving…' : 'Save access'}</Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

// ------------------------------------------------------------------ page

export default function UsersPage() {
  const { user: me } = useAuth();
  const qc = useQueryClient();
  const usersQ = useAdminUsers();
  const sendersQ = useAdminSenders();
  const senders = sendersQ.data ?? [];
  const senderById = useMemo(() => new Map(senders.map(s => [s.id, s])), [senders]);

  const [search, setSearch] = useState('');
  const [createOpen, setCreateOpen] = useState(false);
  const [editing, setEditing] = useState<AdminUser | null>(null);
  const [access, setAccess] = useState<AdminUser | null>(null);
  const [resetting, setResetting] = useState<AdminUser | null>(null);
  const [toggling, setToggling] = useState<AdminUser | null>(null);
  const [revealed, setRevealed] = useState<{ who: string; password: string } | null>(null);

  const users = useMemo(() => {
    const q = search.trim().toLowerCase();
    const all = usersQ.data ?? [];
    return q ? all.filter(u => u.name.toLowerCase().includes(q) || u.email.toLowerCase().includes(q)) : all;
  }, [usersQ.data, search]);

  async function resetPassword(u: AdminUser) {
    try {
      const res = await api.post<{ tempPassword: string }>(`/admin/users/${u.id}/reset-password`);
      qc.invalidateQueries({ queryKey: ADMIN_USERS_KEY });
      setRevealed({ who: u.name, password: res.tempPassword });
    } catch (err) { toast.error(errorMessage(err)); throw err; }
  }

  async function toggleDisabled(u: AdminUser) {
    try {
      await api.patch(`/admin/users/${u.id}`, { disabled: !u.disabled });
      qc.invalidateQueries({ queryKey: ADMIN_USERS_KEY });
      toast.success(u.disabled ? `${u.name} can sign in again.` : `${u.name} has been disabled and signed out.`);
    } catch (err) { toast.error(errorMessage(err)); throw err; }
  }

  return (
    <>
      <PageHeader
        title="Users"
        description="People who can sign in to KIIT Mailer. There is no public sign-up: every account is created here."
        actions={<Button onClick={() => setCreateOpen(true)}><UserPlus /> Add user</Button>}
      />
      <PageBody className="space-y-4">
        {usersQ.isLoading ? <TableSkeleton cols={6} /> : usersQ.isError ? (
          <p role="alert" className="text-sm text-destructive">Could not load users: {errorMessage(usersQ.error)}</p>
        ) : (usersQ.data?.length ?? 0) === 0 ? (
          <EmptyState icon={Users} title="No users yet" action={<Button onClick={() => setCreateOpen(true)}><UserPlus /> Add the first user</Button>}>
            Add the people who will send official mail, then allot them the sender accounts they may use.
          </EmptyState>
        ) : (
          <>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <SearchBox value={search} onChange={setSearch} label="Search users" placeholder="Search by name or email" />
              <span className="text-sm text-muted-foreground">{users.length} of {usersQ.data!.length} users</span>
            </div>
            <div className="overflow-hidden rounded-lg border bg-card">
              <Table>
                <TableHeader>
                  <TableRow className="bg-muted/50 hover:bg-muted/50">
                    <TableHead className="pl-4">Name</TableHead>
                    <TableHead>Role</TableHead>
                    <TableHead>Status</TableHead>
                    <TableHead>Last sign-in</TableHead>
                    <TableHead>Sender accounts</TableHead>
                    <TableHead className="w-12 pr-4"><span className="sr-only">Actions</span></TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {users.length === 0 && (
                    <TableRow><TableCell colSpan={6} className="py-10 text-center text-sm text-muted-foreground">No users match “{search}”.</TableCell></TableRow>
                  )}
                  {users.map(u => {
                    const isSelf = u.id === me?.id;
                    return (
                      <TableRow key={u.id} className={u.disabled ? 'text-muted-foreground' : undefined}>
                        <TableCell className="pl-4">
                          <div className="font-medium text-foreground">
                            {u.name}{isSelf && <span className="ml-2 text-xs font-normal text-muted-foreground">(you)</span>}
                          </div>
                          <div className="text-xs text-muted-foreground">{u.email}</div>
                        </TableCell>
                        <TableCell><RoleBadge role={u.role} /></TableCell>
                        <TableCell><StatusBadge user={u} /></TableCell>
                        <TableCell className="whitespace-nowrap text-sm">
                          {u.lastLoginAt
                            ? <span title={formatDateTime(u.lastLoginAt)}>{relativeTime(u.lastLoginAt)}</span>
                            : <span className="text-muted-foreground">Never</span>}
                        </TableCell>
                        <TableCell className="max-w-[22rem] whitespace-normal">
                          {u.role === 'admin' ? (
                            <span className="text-xs text-muted-foreground">All enabled accounts</span>
                          ) : u.senderAccountIds.length === 0 ? (
                            <button type="button" onClick={() => setAccess(u)} className="text-xs text-muted-foreground underline-offset-2 hover:underline">
                              None allotted
                            </button>
                          ) : (
                            <div className="flex flex-wrap gap-1">
                              {u.senderAccountIds.map(id => {
                                const s = senderById.get(id);
                                return (
                                  <Tooltip key={id}>
                                    <TooltipTrigger asChild>
                                      <span className="inline-flex max-w-[11rem] truncate rounded border bg-secondary px-1.5 py-0.5 text-xs text-secondary-foreground">
                                        {s?.displayName ?? 'Unknown account'}
                                      </span>
                                    </TooltipTrigger>
                                    <TooltipContent>{s ? `${s.email}${s.disabled ? ' (disabled)' : ''}` : id}</TooltipContent>
                                  </Tooltip>
                                );
                              })}
                            </div>
                          )}
                        </TableCell>
                        <TableCell className="pr-4 text-right">
                          <DropdownMenu modal={false}>
                            <DropdownMenuTrigger asChild>
                              <Button variant="ghost" size="icon-sm" aria-label={`Actions for ${u.name}`}><MoreHorizontal /></Button>
                            </DropdownMenuTrigger>
                            <DropdownMenuContent align="end" className="w-56">
                              <DropdownMenuItem onSelect={() => setEditing(u)}><Pencil /> Edit details</DropdownMenuItem>
                              <DropdownMenuItem onSelect={() => setAccess(u)}><KeyRound /> Manage sender access</DropdownMenuItem>
                              <DropdownMenuItem disabled={isSelf} onSelect={() => setResetting(u)}><RotateCcw /> Reset password</DropdownMenuItem>
                              <DropdownMenuSeparator />
                              <DropdownMenuItem
                                disabled={isSelf} onSelect={() => setToggling(u)}
                                className={u.disabled ? undefined : 'text-destructive focus:text-destructive'}
                              >
                                {u.disabled ? <><UserCheck /> Enable access</> : <><UserX /> Disable access</>}
                              </DropdownMenuItem>
                            </DropdownMenuContent>
                          </DropdownMenu>
                        </TableCell>
                      </TableRow>
                    );
                  })}
                </TableBody>
              </Table>
            </div>
          </>
        )}
      </PageBody>

      <CreateUserDialog open={createOpen} onOpenChange={setCreateOpen} senders={senders} />
      {editing && <EditUserDialog key={editing.id} user={editing} isSelf={editing.id === me?.id} onClose={() => setEditing(null)} />}
      {access && <SenderAccessDialog key={access.id} user={access} senders={senders} onClose={() => setAccess(null)} />}

      <ConfirmDialog
        open={!!resetting} onOpenChange={o => { if (!o) setResetting(null); }}
        title={`Reset the password for ${resetting?.name ?? ''}?`}
        description={<>
          <p>A new temporary password will be generated and shown to you once. {resetting?.name} will be signed out everywhere and must choose a new password at their next sign-in.</p>
          <p>Their current password stops working immediately.</p>
        </>}
        confirmLabel="Reset password"
        onConfirm={() => resetPassword(resetting!)}
      />
      <ConfirmDialog
        open={!!toggling} onOpenChange={o => { if (!o) setToggling(null); }}
        title={toggling?.disabled ? `Enable ${toggling?.name}?` : `Disable ${toggling?.name ?? ''}?`}
        description={toggling?.disabled
          ? <p>{toggling?.name} will be able to sign in again with their existing password.</p>
          : <p>{toggling?.name} will be signed out immediately and will not be able to sign in until you enable the account again. Their templates and send history are kept.</p>}
        confirmLabel={toggling?.disabled ? 'Enable access' : 'Disable access'}
        destructive={!toggling?.disabled}
        onConfirm={() => toggleDisabled(toggling!)}
      />

      <Dialog open={!!revealed} onOpenChange={o => { if (!o) setRevealed(null); }}>
        <DialogContent className="sm:max-w-lg">
          <DialogHeader>
            <DialogTitle className="font-serif text-xl">Password reset</DialogTitle>
            <DialogDescription>{revealed?.who} has been signed out. Give them this temporary password to sign in again.</DialogDescription>
          </DialogHeader>
          {revealed && <TempPasswordReveal password={revealed.password} who={revealed.who} />}
          <DialogFooter><Button onClick={() => setRevealed(null)}>Done</Button></DialogFooter>
        </DialogContent>
      </Dialog>
    </>
  );
}
