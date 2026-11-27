CREATE TYPE "public"."editor_mode" AS ENUM('rich', 'html');--> statement-breakpoint
CREATE TYPE "public"."job_mode" AS ENUM('personal', 'bcc_broadcast');--> statement-breakpoint
CREATE TYPE "public"."job_status" AS ENUM('queued', 'sending', 'paused', 'completed', 'cancelled');--> statement-breakpoint
CREATE TYPE "public"."message_status" AS ENUM('queued', 'sending', 'sent', 'failed', 'skipped');--> statement-breakpoint
CREATE TYPE "public"."user_role" AS ENUM('admin', 'sender');--> statement-breakpoint
CREATE TYPE "public"."suppression_reason" AS ENUM('unsubscribed', 'bounced', 'manual');--> statement-breakpoint
CREATE TYPE "public"."upload_kind" AS ENUM('attachment', 'image');--> statement-breakpoint
CREATE TYPE "public"."template_visibility" AS ENUM('private', 'shared_all', 'shared_users');--> statement-breakpoint
CREATE TABLE "audit_log" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"actor_id" uuid,
	"actor_email" text,
	"action" text NOT NULL,
	"target_type" text,
	"target_id" text,
	"meta" jsonb,
	"ip" text,
	"at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "daily_usage" (
	"sender_account_id" uuid NOT NULL,
	"day" date NOT NULL,
	"recipients_sent" integer DEFAULT 0 NOT NULL,
	CONSTRAINT "daily_usage_sender_account_id_day_pk" PRIMARY KEY("sender_account_id","day")
);
--> statement-breakpoint
CREATE TABLE "job_attachments" (
	"job_id" uuid NOT NULL,
	"upload_id" uuid NOT NULL,
	CONSTRAINT "job_attachments_job_id_upload_id_pk" PRIMARY KEY("job_id","upload_id")
);
--> statement-breakpoint
CREATE TABLE "login_attempts" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"email" text NOT NULL,
	"ip" text,
	"success" boolean NOT NULL,
	"at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "send_jobs" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"user_id" uuid NOT NULL,
	"sender_account_id" uuid NOT NULL,
	"template_id" uuid,
	"mode" "job_mode" DEFAULT 'personal' NOT NULL,
	"is_test" boolean DEFAULT false NOT NULL,
	"subject" text NOT NULL,
	"preheader" text DEFAULT '' NOT NULL,
	"editor_mode" "editor_mode" NOT NULL,
	"body_html" text NOT NULL,
	"body_text" text,
	"theme" text NOT NULL,
	"include_signature" boolean DEFAULT true NOT NULL,
	"global_cc" text,
	"global_bcc" text,
	"unsubscribe_enabled" boolean DEFAULT false NOT NULL,
	"status" "job_status" DEFAULT 'queued' NOT NULL,
	"total_messages" integer DEFAULT 0 NOT NULL,
	"total_recipients" integer DEFAULT 0 NOT NULL,
	"sent_count" integer DEFAULT 0 NOT NULL,
	"failed_count" integer DEFAULT 0 NOT NULL,
	"skipped_count" integer DEFAULT 0 NOT NULL,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL,
	"started_at" timestamp with time zone,
	"finished_at" timestamp with time zone,
	"scheduled_at" timestamp with time zone
);
--> statement-breakpoint
CREATE TABLE "send_messages" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"job_id" uuid NOT NULL,
	"seq" integer NOT NULL,
	"to_addr" text NOT NULL,
	"cc" text,
	"bcc" text,
	"recipient_count" integer DEFAULT 1 NOT NULL,
	"row_data" jsonb DEFAULT '{}'::jsonb NOT NULL,
	"status" "message_status" DEFAULT 'queued' NOT NULL,
	"attempts" integer DEFAULT 0 NOT NULL,
	"error" text,
	"retryable" boolean,
	"smtp_message_id" text,
	"claimed_at" timestamp with time zone,
	"sent_at" timestamp with time zone
);
--> statement-breakpoint
CREATE TABLE "sender_accounts" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"email" text NOT NULL,
	"display_name" text NOT NULL,
	"reply_to" text,
	"app_password_enc" text NOT NULL,
	"app_password_iv" text NOT NULL,
	"app_password_tag" text NOT NULL,
	"key_version" integer DEFAULT 1 NOT NULL,
	"app_password_hint" text,
	"daily_limit" integer DEFAULT 2000 NOT NULL,
	"signature_doc" jsonb,
	"signature_html" text,
	"default_theme" text DEFAULT 'heritage' NOT NULL,
	"last_verified_at" timestamp with time zone,
	"last_verify_error" text,
	"disabled" boolean DEFAULT false NOT NULL,
	"created_by" uuid,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL,
	"updated_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "sessions" (
	"token_hash" text PRIMARY KEY NOT NULL,
	"user_id" uuid NOT NULL,
	"expires_at" timestamp with time zone NOT NULL,
	"last_seen_at" timestamp with time zone DEFAULT now() NOT NULL,
	"ip" text,
	"user_agent" text,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "suppressions" (
	"email" text PRIMARY KEY NOT NULL,
	"reason" "suppression_reason" NOT NULL,
	"note" text,
	"sender_account_id" uuid,
	"created_by" uuid,
	"at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "template_shares" (
	"template_id" uuid NOT NULL,
	"user_id" uuid NOT NULL,
	CONSTRAINT "template_shares_template_id_user_id_pk" PRIMARY KEY("template_id","user_id")
);
--> statement-breakpoint
CREATE TABLE "templates" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"owner_id" uuid NOT NULL,
	"name" text NOT NULL,
	"category" text,
	"subject" text DEFAULT '' NOT NULL,
	"preheader" text DEFAULT '' NOT NULL,
	"editor_mode" "editor_mode" DEFAULT 'rich' NOT NULL,
	"doc_json" jsonb,
	"html" text DEFAULT '' NOT NULL,
	"text" text,
	"theme" text DEFAULT 'heritage' NOT NULL,
	"include_signature" boolean DEFAULT true NOT NULL,
	"default_sender_account_id" uuid,
	"visibility" "template_visibility" DEFAULT 'private' NOT NULL,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL,
	"updated_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "uploads" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"owner_id" uuid NOT NULL,
	"kind" "upload_kind" NOT NULL,
	"provider" text NOT NULL,
	"url" text,
	"local_id" text,
	"filename" text NOT NULL,
	"content_type" text NOT NULL,
	"size" bigint NOT NULL,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "user_sender_access" (
	"user_id" uuid NOT NULL,
	"sender_account_id" uuid NOT NULL,
	"granted_by" uuid,
	"granted_at" timestamp with time zone DEFAULT now() NOT NULL,
	CONSTRAINT "user_sender_access_user_id_sender_account_id_pk" PRIMARY KEY("user_id","sender_account_id")
);
--> statement-breakpoint
CREATE TABLE "users" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"email" text NOT NULL,
	"name" text NOT NULL,
	"password_hash" text NOT NULL,
	"role" "user_role" DEFAULT 'sender' NOT NULL,
	"must_change_password" boolean DEFAULT true NOT NULL,
	"disabled" boolean DEFAULT false NOT NULL,
	"last_login_at" timestamp with time zone,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL,
	"updated_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
ALTER TABLE "audit_log" ADD CONSTRAINT "audit_log_actor_id_users_id_fk" FOREIGN KEY ("actor_id") REFERENCES "public"."users"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "daily_usage" ADD CONSTRAINT "daily_usage_sender_account_id_sender_accounts_id_fk" FOREIGN KEY ("sender_account_id") REFERENCES "public"."sender_accounts"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "job_attachments" ADD CONSTRAINT "job_attachments_job_id_send_jobs_id_fk" FOREIGN KEY ("job_id") REFERENCES "public"."send_jobs"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "job_attachments" ADD CONSTRAINT "job_attachments_upload_id_uploads_id_fk" FOREIGN KEY ("upload_id") REFERENCES "public"."uploads"("id") ON DELETE restrict ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "send_jobs" ADD CONSTRAINT "send_jobs_user_id_users_id_fk" FOREIGN KEY ("user_id") REFERENCES "public"."users"("id") ON DELETE restrict ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "send_jobs" ADD CONSTRAINT "send_jobs_sender_account_id_sender_accounts_id_fk" FOREIGN KEY ("sender_account_id") REFERENCES "public"."sender_accounts"("id") ON DELETE restrict ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "send_jobs" ADD CONSTRAINT "send_jobs_template_id_templates_id_fk" FOREIGN KEY ("template_id") REFERENCES "public"."templates"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "send_messages" ADD CONSTRAINT "send_messages_job_id_send_jobs_id_fk" FOREIGN KEY ("job_id") REFERENCES "public"."send_jobs"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "sender_accounts" ADD CONSTRAINT "sender_accounts_created_by_users_id_fk" FOREIGN KEY ("created_by") REFERENCES "public"."users"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "sessions" ADD CONSTRAINT "sessions_user_id_users_id_fk" FOREIGN KEY ("user_id") REFERENCES "public"."users"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "suppressions" ADD CONSTRAINT "suppressions_sender_account_id_sender_accounts_id_fk" FOREIGN KEY ("sender_account_id") REFERENCES "public"."sender_accounts"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "suppressions" ADD CONSTRAINT "suppressions_created_by_users_id_fk" FOREIGN KEY ("created_by") REFERENCES "public"."users"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "template_shares" ADD CONSTRAINT "template_shares_template_id_templates_id_fk" FOREIGN KEY ("template_id") REFERENCES "public"."templates"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "template_shares" ADD CONSTRAINT "template_shares_user_id_users_id_fk" FOREIGN KEY ("user_id") REFERENCES "public"."users"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "templates" ADD CONSTRAINT "templates_owner_id_users_id_fk" FOREIGN KEY ("owner_id") REFERENCES "public"."users"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "templates" ADD CONSTRAINT "templates_default_sender_account_id_sender_accounts_id_fk" FOREIGN KEY ("default_sender_account_id") REFERENCES "public"."sender_accounts"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "uploads" ADD CONSTRAINT "uploads_owner_id_users_id_fk" FOREIGN KEY ("owner_id") REFERENCES "public"."users"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "user_sender_access" ADD CONSTRAINT "user_sender_access_user_id_users_id_fk" FOREIGN KEY ("user_id") REFERENCES "public"."users"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "user_sender_access" ADD CONSTRAINT "user_sender_access_sender_account_id_sender_accounts_id_fk" FOREIGN KEY ("sender_account_id") REFERENCES "public"."sender_accounts"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "user_sender_access" ADD CONSTRAINT "user_sender_access_granted_by_users_id_fk" FOREIGN KEY ("granted_by") REFERENCES "public"."users"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
CREATE INDEX "audit_log_at_idx" ON "audit_log" USING btree ("at");--> statement-breakpoint
CREATE INDEX "audit_log_actor_idx" ON "audit_log" USING btree ("actor_id","at");--> statement-breakpoint
CREATE INDEX "login_attempts_email_at_idx" ON "login_attempts" USING btree ("email","at");--> statement-breakpoint
CREATE INDEX "send_jobs_user_idx" ON "send_jobs" USING btree ("user_id","created_at");--> statement-breakpoint
CREATE INDEX "send_jobs_status_idx" ON "send_jobs" USING btree ("status");--> statement-breakpoint
CREATE INDEX "send_messages_job_status_idx" ON "send_messages" USING btree ("job_id","status","seq");--> statement-breakpoint
CREATE UNIQUE INDEX "sender_accounts_email_uq" ON "sender_accounts" USING btree ("email");--> statement-breakpoint
CREATE INDEX "sessions_user_idx" ON "sessions" USING btree ("user_id");--> statement-breakpoint
CREATE INDEX "templates_owner_idx" ON "templates" USING btree ("owner_id");--> statement-breakpoint
CREATE INDEX "uploads_owner_idx" ON "uploads" USING btree ("owner_id");--> statement-breakpoint
CREATE UNIQUE INDEX "users_email_uq" ON "users" USING btree ("email");