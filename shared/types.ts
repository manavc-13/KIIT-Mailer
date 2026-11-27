// Shared API contracts between the React client (src/) and the Hono API (server/).
// Every endpoint's request/response shape lives here. Change a contract here
// first, then both sides follow. All routes are mounted under /api.

export type Role = 'admin' | 'sender';
export type ThemeId = 'heritage' | 'meridian' | 'letterhead' | 'slate' | 'plain';
export type EditorMode = 'rich' | 'html';
export type JobMode = 'personal' | 'bcc_broadcast';
export type JobStatus = 'queued' | 'sending' | 'paused' | 'completed' | 'cancelled';
export type MessageStatus = 'queued' | 'sending' | 'sent' | 'failed' | 'skipped';
export type TemplateVisibility = 'private' | 'shared_all' | 'shared_users';

/** Error body for every non-2xx response: `{ error: string, code?: string, details?: unknown }`. */
export interface ApiErrorBody { error: string; code?: string; details?: unknown }

// ------------------------------------------------------------------ auth
// POST /api/auth/login            LoginRequest -> MeResponse   (sets httpOnly session cookie)
// POST /api/auth/logout           -> { ok: true }
// GET  /api/auth/me               -> MeResponse (401 if no session)
// POST /api/auth/change-password  ChangePasswordRequest -> { ok: true } (other sessions revoked)
export interface LoginRequest { email: string; password: string }
export interface ChangePasswordRequest { currentPassword: string; newPassword: string }
export interface PublicUser {
  id: string; email: string; name: string; role: Role;
  mustChangePassword: boolean; disabled: boolean;
  lastLoginAt: string | null; createdAt: string;
}
export interface MeResponse { user: PublicUser }

// ------------------------------------------------------------------ sender accounts
// GET  /api/senders                    -> { senders: SenderAccount[] }  (only accounts allotted to the caller; admins see all enabled)
// --- admin only ---
// GET    /api/admin/senders            -> { senders: SenderAccountAdmin[] }
// POST   /api/admin/senders            CreateSenderRequest -> { sender: SenderAccountAdmin }  (SMTP verify() must pass first)
// PATCH  /api/admin/senders/:id        UpdateSenderRequest -> { sender: SenderAccountAdmin }
// POST   /api/admin/senders/:id/verify -> { ok: boolean, error?: string, sender: SenderAccountAdmin }
// POST   /api/admin/senders/test-credentials  { email, appPassword } -> { ok: boolean, error?: string }
export interface SenderAccount {
  id: string; email: string; displayName: string; replyTo: string | null;
  defaultTheme: ThemeId; signatureHtml: string | null; signatureDoc: unknown | null;
  dailyLimit: number; usedToday: number; disabled: boolean;
}
export interface SenderAccountAdmin extends SenderAccount {
  appPasswordHint: string | null; // e.g. "abcd" -> UI shows ••••abcd
  lastVerifiedAt: string | null; lastVerifyError: string | null;
  createdAt: string; userIds: string[]; // users allotted to this account
}
export interface CreateSenderRequest {
  email: string; displayName: string; replyTo?: string | null; appPassword: string;
  dailyLimit?: number; defaultTheme?: ThemeId; signatureHtml?: string | null; signatureDoc?: unknown | null;
}
export interface UpdateSenderRequest extends Partial<Omit<CreateSenderRequest, 'email'>> { disabled?: boolean }

// ------------------------------------------------------------------ users (admin)
// GET    /api/admin/users                    -> { users: AdminUser[] }
// POST   /api/admin/users                    CreateUserRequest -> { user: AdminUser, tempPassword: string }
// PATCH  /api/admin/users/:id                UpdateUserRequest -> { user: AdminUser }   (disabling revokes sessions)
// POST   /api/admin/users/:id/reset-password -> { tempPassword: string }                (revokes sessions, forces change)
// PUT    /api/admin/users/:id/senders        { senderAccountIds: string[] } -> { user: AdminUser }
export interface AdminUser extends PublicUser { senderAccountIds: string[] }
export interface CreateUserRequest { email: string; name: string; role: Role; senderAccountIds?: string[] }
export interface UpdateUserRequest { name?: string; role?: Role; disabled?: boolean }

// ------------------------------------------------------------------ audit (admin)
// GET /api/admin/audit?actorId=&action=&from=&to=&cursor= -> { entries: AuditEntry[], nextCursor: string | null }
export type AuditAction =
  | 'auth.login' | 'auth.login_failed' | 'auth.logout' | 'auth.password_changed'
  | 'user.create' | 'user.update' | 'user.disable' | 'user.enable' | 'user.reset_password' | 'user.senders_set'
  | 'sender.create' | 'sender.update' | 'sender.password_rotated' | 'sender.verify' | 'sender.disable' | 'sender.enable'
  | 'template.create' | 'template.update' | 'template.delete' | 'template.share'
  | 'job.create' | 'job.start' | 'job.pause' | 'job.cancel' | 'job.complete' | 'job.retry_failed' | 'job.test_send'
  | 'suppression.add' | 'suppression.remove';
export interface AuditEntry {
  id: string; actorId: string | null; actorEmail: string | null; action: AuditAction | string;
  targetType: string | null; targetId: string | null; meta: Record<string, unknown> | null; ip: string | null; at: string;
}

// ------------------------------------------------------------------ templates
// GET    /api/templates               -> { templates: TemplateSummary[] }  (own + shared with me)
// GET    /api/templates/:id           -> { template: Template }
// POST   /api/templates               SaveTemplateRequest -> { template: Template }
// PUT    /api/templates/:id           SaveTemplateRequest -> { template: Template }   (owner only)
// DELETE /api/templates/:id           -> { ok: true }                                 (owner or admin)
// POST   /api/templates/:id/duplicate -> { template: Template }                       (copy becomes caller's private template)
// PUT    /api/templates/:id/share     { visibility, userIds?: string[] } -> { template: Template }  (admin only)
// POST   /api/templates/import        { templates: SaveTemplateRequest[] } -> { created: number }
export interface TemplateSummary {
  id: string; name: string; category: string | null; subject: string; theme: ThemeId;
  editorMode: EditorMode; visibility: TemplateVisibility; ownerId: string; ownerName: string;
  isOwner: boolean; defaultSenderAccountId: string | null; updatedAt: string;
}
export interface Template extends TemplateSummary {
  preheader: string; docJson: unknown | null; html: string; text: string | null;
  includeSignature: boolean; sharedUserIds: string[];
}
export interface SaveTemplateRequest {
  name: string; category?: string | null; subject: string; preheader?: string;
  editorMode: EditorMode; docJson?: unknown | null; html: string; text?: string | null;
  theme: ThemeId; includeSignature?: boolean; defaultSenderAccountId?: string | null;
}

// ------------------------------------------------------------------ uploads
// Attachments (private) and in-body images (public, hosted so no data: URIs are ever sent).
// GET  /api/uploads/config                 -> UploadConfig
// POST /api/uploads/blob-token             (Vercel Blob client-upload handshake; body is the @vercel/blob/client payload)
// POST /api/uploads/register               RegisterUploadRequest -> { upload: UploadRef }   (after a blob upload completes)
// POST /api/uploads/local?kind=&filename=&contentType=  raw body -> { upload: UploadRef }   (local dev provider only)
// POST /api/uploads/image-from-url         { url } -> { upload: UploadRef }   (rehost pasted Google Docs images, server-side fetch)
// GET  /api/uploads/:id/content            streams the file (owner only; local provider or private blob)
// DELETE /api/uploads/:id                  -> { ok: true } (owner only)
export type UploadKind = 'attachment' | 'image';
export interface UploadConfig {
  provider: 'blob' | 'local' | 'inline';
  publicImages: boolean; // a public Blob store is configured for in-body images
  maxRawAttachmentBytes: number; maxEncodedMessageBytes: number; softWarnBytes: number; mimeOverheadBytes: number;
  degradedReason: string | null;
}
export interface UploadRef {
  id: string; kind: UploadKind; filename: string; contentType: string; size: number;
  /** For images: the URL to put in <img src>. Absolute https URL for blob, /api/uploads/:id/content for local dev. */
  publicUrl: string | null;
}
export interface RegisterUploadRequest { kind: UploadKind; url: string; filename: string; contentType: string; size: number }

// ------------------------------------------------------------------ render / preview
// POST /api/render/preview  RenderRequest -> RenderResult
// Uses EXACTLY the same server pipeline as sending (server/lib/render/index.ts).
export interface ComposeContent {
  subject: string; preheader: string; editorMode: EditorMode;
  /** rich mode: email-safe BODY fragment from the editor serializer. html mode: a full HTML document. */
  bodyHtml: string;
  /** Optional hand-written plain-text part; when empty the server generates one from the HTML. */
  bodyText?: string | null;
  theme: ThemeId; includeSignature: boolean;
}
export interface RenderRequest extends ComposeContent {
  senderAccountId: string;
  row?: Record<string, string>; // placeholder values for the previewed recipient
  unsubscribe?: boolean;
}
export interface LintIssue {
  level: 'error' | 'warn' | 'info';
  code: string; // e.g. 'spam_word', 'gmail_clip', 'data_uri_image', 'missing_alt', 'http_link', 'shortener', 'link_mismatch', 'caps_subject'
  message: string;
}
export interface RenderResult {
  subject: string; html: string; text: string; sizeBytes: number; lint: LintIssue[];
  /** Fonts that will fall back in Gmail/Outlook (Enhanced web fonts used in the HTML). */
  webFontsUsed: string[];
}

// ------------------------------------------------------------------ jobs (sending)
// POST /api/jobs                 CreateJobRequest -> { job: Job, warnings: string[] }
// POST /api/jobs/test            TestSendRequest -> { results: { to: string; success: boolean; error?: string }[] }
// POST /api/jobs/:id/run         -> RunJobResponse   (client calls repeatedly until done; each call claims messages for ~32 s, well inside the 60 s function limit)
// POST /api/jobs/:id/pause       -> { job: Job }
// POST /api/jobs/:id/cancel      -> { job: Job }
// POST /api/jobs/:id/retry-failed -> { job: Job }    (re-queues failed messages)
// GET  /api/jobs?scope=mine|all&cursor= -> { jobs: Job[], nextCursor: string | null }   (scope=all admin only)
// GET  /api/jobs/active          -> { jobs: Job[] }   (caller's queued/sending/paused jobs, for the "reopen" banner)
// GET  /api/jobs/:id             -> { job: Job, messages: JobMessage[] }
// GET  /api/jobs/:id/export.csv  -> text/csv
export interface JobRecipient {
  to: string;            // one or more addresses, comma-separated (personal mode: usually one)
  cc?: string; bcc?: string;
  data?: Record<string, string>; // placeholder values (row) for this message
}
export interface CreateJobRequest extends ComposeContent {
  senderAccountId: string; templateId?: string | null; mode: JobMode;
  /** personal mode: one entry per message. bcc_broadcast: each entry's `to` is a BCC address; server chunks <= 90 per message. */
  recipients: JobRecipient[];
  globalCc?: string; globalBcc?: string;
  attachmentIds: string[]; // UploadRef ids (kind 'attachment')
  /** 'auto' = on when > 50 recipients outside kiit.ac.in */
  unsubscribe: boolean | 'auto';
  dedupe?: boolean;
}
export interface TestSendRequest extends ComposeContent {
  senderAccountId: string; to: string[]; row?: Record<string, string>; attachmentIds: string[];
}
export interface Job {
  id: string; userId: string; userName?: string; senderAccountId: string; senderEmail: string;
  templateId: string | null; mode: JobMode; subject: string; status: JobStatus;
  totalMessages: number; totalRecipients: number; sentCount: number; failedCount: number; skippedCount: number;
  createdAt: string; startedAt: string | null; finishedAt: string | null;
}
export interface JobMessage {
  id: string; seq: number; to: string; cc: string | null; bcc: string | null; recipientCount: number;
  status: MessageStatus; attempts: number; error: string | null; retryable: boolean | null; sentAt: string | null;
}
/**
 * `done: true` means "stop calling /run". That includes a pause (job.status === 'paused'),
 * so check job.status to tell paused apart from completed/cancelled. `error` explains an automatic
 * pause (App Password rejected, daily quota exhausted, sender disabled, attachment missing).
 */
export interface RunJobResponse { job: Job; processed: JobMessage[]; done: boolean; quotaExhausted?: boolean; error?: string }

// ------------------------------------------------------------------ suppressions (admin)
// GET    /api/admin/suppressions -> { suppressions: { email, reason, note, at }[] }
// POST   /api/admin/suppressions { email, note? } -> { ok: true }
// DELETE /api/admin/suppressions/:email -> { ok: true }
// Public, no auth: GET/POST /u/:token  (one-click unsubscribe; RFC 8058)
