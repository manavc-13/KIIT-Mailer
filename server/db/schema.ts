// KIIT Mailer database schema (Drizzle ORM, Neon Postgres).
// Source of truth for every table. Change it here, then `npm run db:generate`
// and `npm run db:migrate`.
import {
  pgTable, pgEnum, uuid, text, boolean, timestamp, integer, jsonb, primaryKey, index, uniqueIndex, date, bigint,
} from 'drizzle-orm/pg-core';

export const roleEnum = pgEnum('user_role', ['admin', 'sender']);
export const visibilityEnum = pgEnum('template_visibility', ['private', 'shared_all', 'shared_users']);
export const editorModeEnum = pgEnum('editor_mode', ['rich', 'html']);
export const jobModeEnum = pgEnum('job_mode', ['personal', 'bcc_broadcast']);
export const jobStatusEnum = pgEnum('job_status', ['queued', 'sending', 'paused', 'completed', 'cancelled']);
export const messageStatusEnum = pgEnum('message_status', ['queued', 'sending', 'sent', 'failed', 'skipped']);
export const suppressionReasonEnum = pgEnum('suppression_reason', ['unsubscribed', 'bounced', 'manual']);
export const uploadKindEnum = pgEnum('upload_kind', ['attachment', 'image']);

const ts = (name: string) => timestamp(name, { withTimezone: true });

// ---------------------------------------------------------------- users
export const users = pgTable('users', {
  id: uuid('id').primaryKey().defaultRandom(),
  email: text('email').notNull(),
  name: text('name').notNull(),
  passwordHash: text('password_hash').notNull(),
  role: roleEnum('role').notNull().default('sender'),
  mustChangePassword: boolean('must_change_password').notNull().default(true),
  disabled: boolean('disabled').notNull().default(false),
  lastLoginAt: ts('last_login_at'),
  createdAt: ts('created_at').notNull().defaultNow(),
  updatedAt: ts('updated_at').notNull().defaultNow(),
}, (t) => [uniqueIndex('users_email_uq').on(t.email)]);

export const sessions = pgTable('sessions', {
  tokenHash: text('token_hash').primaryKey(), // sha256(hex) of the cookie token; the raw token is never stored
  userId: uuid('user_id').notNull().references(() => users.id, { onDelete: 'cascade' }),
  expiresAt: ts('expires_at').notNull(),
  lastSeenAt: ts('last_seen_at').notNull().defaultNow(),
  ip: text('ip'),
  userAgent: text('user_agent'),
  createdAt: ts('created_at').notNull().defaultNow(),
}, (t) => [index('sessions_user_idx').on(t.userId)]);

export const loginAttempts = pgTable('login_attempts', {
  id: uuid('id').primaryKey().defaultRandom(),
  email: text('email').notNull(),
  ip: text('ip'),
  success: boolean('success').notNull(),
  at: ts('at').notNull().defaultNow(),
}, (t) => [index('login_attempts_email_at_idx').on(t.email, t.at)]);

// ---------------------------------------------------------------- sender accounts
export const senderAccounts = pgTable('sender_accounts', {
  id: uuid('id').primaryKey().defaultRandom(),
  email: text('email').notNull(), // always @kiit.ac.in
  displayName: text('display_name').notNull(),
  replyTo: text('reply_to'),
  // AES-256-GCM encrypted Google App Password (server/lib/crypto.ts). Never returned by any API.
  appPasswordEnc: text('app_password_enc').notNull(),
  appPasswordIv: text('app_password_iv').notNull(),
  appPasswordTag: text('app_password_tag').notNull(),
  keyVersion: integer('key_version').notNull().default(1),
  appPasswordHint: text('app_password_hint'), // last 4 chars, for "••••abcd" display
  dailyLimit: integer('daily_limit').notNull().default(2000),
  signatureDoc: jsonb('signature_doc'), // TipTap JSON
  signatureHtml: text('signature_html'), // compiled email-safe HTML
  defaultTheme: text('default_theme').notNull().default('heritage'), // shared/themes id
  lastVerifiedAt: ts('last_verified_at'),
  lastVerifyError: text('last_verify_error'),
  disabled: boolean('disabled').notNull().default(false),
  createdBy: uuid('created_by').references(() => users.id, { onDelete: 'set null' }),
  createdAt: ts('created_at').notNull().defaultNow(),
  updatedAt: ts('updated_at').notNull().defaultNow(),
}, (t) => [uniqueIndex('sender_accounts_email_uq').on(t.email)]);

export const userSenderAccess = pgTable('user_sender_access', {
  userId: uuid('user_id').notNull().references(() => users.id, { onDelete: 'cascade' }),
  senderAccountId: uuid('sender_account_id').notNull().references(() => senderAccounts.id, { onDelete: 'cascade' }),
  grantedBy: uuid('granted_by').references(() => users.id, { onDelete: 'set null' }),
  grantedAt: ts('granted_at').notNull().defaultNow(),
}, (t) => [primaryKey({ columns: [t.userId, t.senderAccountId] })]);

// ---------------------------------------------------------------- templates
export const templates = pgTable('templates', {
  id: uuid('id').primaryKey().defaultRandom(),
  ownerId: uuid('owner_id').notNull().references(() => users.id, { onDelete: 'cascade' }),
  name: text('name').notNull(),
  category: text('category'),
  subject: text('subject').notNull().default(''),
  preheader: text('preheader').notNull().default(''),
  editorMode: editorModeEnum('editor_mode').notNull().default('rich'),
  docJson: jsonb('doc_json'), // TipTap JSON (rich mode source of truth)
  html: text('html').notNull().default(''), // compiled body HTML (rich) or full HTML (html mode)
  text: text('text'), // optional hand-written plain-text part
  theme: text('theme').notNull().default('heritage'),
  includeSignature: boolean('include_signature').notNull().default(true),
  defaultSenderAccountId: uuid('default_sender_account_id').references(() => senderAccounts.id, { onDelete: 'set null' }),
  visibility: visibilityEnum('visibility').notNull().default('private'),
  createdAt: ts('created_at').notNull().defaultNow(),
  updatedAt: ts('updated_at').notNull().defaultNow(),
}, (t) => [index('templates_owner_idx').on(t.ownerId)]);

export const templateShares = pgTable('template_shares', {
  templateId: uuid('template_id').notNull().references(() => templates.id, { onDelete: 'cascade' }),
  userId: uuid('user_id').notNull().references(() => users.id, { onDelete: 'cascade' }),
}, (t) => [primaryKey({ columns: [t.templateId, t.userId] })]);

// ---------------------------------------------------------------- uploads (attachments + in-body images)
export const uploads = pgTable('uploads', {
  id: uuid('id').primaryKey().defaultRandom(),
  ownerId: uuid('owner_id').notNull().references(() => users.id, { onDelete: 'cascade' }),
  kind: uploadKindEnum('kind').notNull(),
  provider: text('provider').notNull(), // 'blob' | 'local' | 'inline'
  url: text('url'), // blob url (private for attachments, public for images)
  localId: text('local_id'), // local dev server temp id
  filename: text('filename').notNull(),
  contentType: text('content_type').notNull(),
  size: bigint('size', { mode: 'number' }).notNull(),
  createdAt: ts('created_at').notNull().defaultNow(),
}, (t) => [index('uploads_owner_idx').on(t.ownerId)]);

// ---------------------------------------------------------------- send jobs
export const sendJobs = pgTable('send_jobs', {
  id: uuid('id').primaryKey().defaultRandom(),
  userId: uuid('user_id').notNull().references(() => users.id, { onDelete: 'restrict' }),
  senderAccountId: uuid('sender_account_id').notNull().references(() => senderAccounts.id, { onDelete: 'restrict' }),
  templateId: uuid('template_id').references(() => templates.id, { onDelete: 'set null' }),
  mode: jobModeEnum('mode').notNull().default('personal'),
  isTest: boolean('is_test').notNull().default(false),
  // Snapshot of what is being sent, so later template edits never change an in-flight job.
  subject: text('subject').notNull(),
  preheader: text('preheader').notNull().default(''),
  editorMode: editorModeEnum('editor_mode').notNull(),
  bodyHtml: text('body_html').notNull(),
  bodyText: text('body_text'),
  theme: text('theme').notNull(),
  includeSignature: boolean('include_signature').notNull().default(true),
  globalCc: text('global_cc'),
  globalBcc: text('global_bcc'),
  unsubscribeEnabled: boolean('unsubscribe_enabled').notNull().default(false),
  status: jobStatusEnum('status').notNull().default('queued'),
  totalMessages: integer('total_messages').notNull().default(0),
  totalRecipients: integer('total_recipients').notNull().default(0), // To+CC+BCC across all messages
  sentCount: integer('sent_count').notNull().default(0),
  failedCount: integer('failed_count').notNull().default(0),
  skippedCount: integer('skipped_count').notNull().default(0),
  createdAt: ts('created_at').notNull().defaultNow(),
  startedAt: ts('started_at'),
  finishedAt: ts('finished_at'),
  // Reserved for scheduled sending (not in v2): when set, a cron runner picks the job up.
  scheduledAt: ts('scheduled_at'),
}, (t) => [index('send_jobs_user_idx').on(t.userId, t.createdAt), index('send_jobs_status_idx').on(t.status)]);

export const sendMessages = pgTable('send_messages', {
  id: uuid('id').primaryKey().defaultRandom(),
  jobId: uuid('job_id').notNull().references(() => sendJobs.id, { onDelete: 'cascade' }),
  seq: integer('seq').notNull(),
  to: text('to_addr').notNull(), // comma-separated
  cc: text('cc'),
  bcc: text('bcc'),
  recipientCount: integer('recipient_count').notNull().default(1),
  rowData: jsonb('row_data').notNull().default({}), // placeholder values for this message
  status: messageStatusEnum('status').notNull().default('queued'),
  attempts: integer('attempts').notNull().default(0),
  error: text('error'),
  retryable: boolean('retryable'),
  smtpMessageId: text('smtp_message_id'),
  claimedAt: ts('claimed_at'),
  sentAt: ts('sent_at'),
}, (t) => [index('send_messages_job_status_idx').on(t.jobId, t.status, t.seq)]);

export const jobAttachments = pgTable('job_attachments', {
  jobId: uuid('job_id').notNull().references(() => sendJobs.id, { onDelete: 'cascade' }),
  uploadId: uuid('upload_id').notNull().references(() => uploads.id, { onDelete: 'restrict' }),
}, (t) => [primaryKey({ columns: [t.jobId, t.uploadId] })]);

// Per sender account, per calendar day (Asia/Kolkata): recipients counted the way Gmail counts them (To+CC+BCC).
export const dailyUsage = pgTable('daily_usage', {
  senderAccountId: uuid('sender_account_id').notNull().references(() => senderAccounts.id, { onDelete: 'cascade' }),
  day: date('day').notNull(),
  recipientsSent: integer('recipients_sent').notNull().default(0),
}, (t) => [primaryKey({ columns: [t.senderAccountId, t.day] })]);

export const suppressions = pgTable('suppressions', {
  email: text('email').primaryKey(), // lower-cased
  reason: suppressionReasonEnum('reason').notNull(),
  note: text('note'),
  senderAccountId: uuid('sender_account_id').references(() => senderAccounts.id, { onDelete: 'set null' }),
  createdBy: uuid('created_by').references(() => users.id, { onDelete: 'set null' }),
  at: ts('at').notNull().defaultNow(),
});

// ---------------------------------------------------------------- audit
export const auditLog = pgTable('audit_log', {
  id: uuid('id').primaryKey().defaultRandom(),
  actorId: uuid('actor_id').references(() => users.id, { onDelete: 'set null' }),
  actorEmail: text('actor_email'), // denormalised so the log survives user deletion and records failed logins
  action: text('action').notNull(), // dotted verb, e.g. 'auth.login', 'sender.create' (see shared/types.ts AuditAction)
  targetType: text('target_type'),
  targetId: text('target_id'),
  meta: jsonb('meta'),
  ip: text('ip'),
  at: ts('at').notNull().defaultNow(),
}, (t) => [index('audit_log_at_idx').on(t.at), index('audit_log_actor_idx').on(t.actorId, t.at)]);
