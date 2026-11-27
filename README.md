# KIIT Mailer

The institutional mailing service of the Internal Quality Assurance Cell (IQAC), KIIT. Staff sign in, pick an
`@kiit.ac.in` sender account the administrator has allotted to them, and send formal, letterhead-styled email
to one person or thousands, personalised per recipient or as a single BCC broadcast.

- **Login wall + access control.** No self-registration. The administrator creates users, assigns roles, and
  allots which sender accounts each user may send from.
- **Sender accounts.** Each `@kiit.ac.in` account has its Google App Password stored **encrypted on the server**
  (AES-256-GCM). It is tested against Gmail before it can be saved, and never sent back to any browser. Each
  account has its own display name, reply-to, signature, default letterhead and daily quota.
- **Templates.** Personal templates, plus templates the administrator shares with everyone or with chosen people.
  Switch template and sender in one click; the sender's signature and letterhead follow automatically.
- **Letterheads.** Five themes: Heritage, Meridian, Letterhead, Slate (ported from the IQAC email designs) and Plain.
- **Editor.** Rich text with email-safe fonts (universal fonts plus web fonts with matched fallbacks), sizes,
  colours, tables, images, buttons, callouts and detail panels. Pasting from Google Docs or Word keeps fonts,
  sizes, colours, lists and tables. An HTML mode is available for hand-built emails.
- **Honest preview.** Gmail (desktop and mobile), Apple Mail on iPhone, an Outlook for Windows approximation,
  dark mode and plain text, all rendered by the same server code that sends. It also shows which fonts fall back
  where, Gmail's 102 KB clipping limit, and a compatibility report.
- **Deliverability.** Every message is multipart (HTML plus plain text) with a proper Message-ID, CSS inlined,
  no `data:` images, paced sending, and per-account quotas counted the way Gmail counts them (To + CC + BCC).
  Optional one-click unsubscribe (RFC 8058) feeds a suppression list. A pre-send checklist flags spam triggers,
  unresolved placeholders, missing alt text, risky links and oversized HTML.
- **Server-side send jobs.** Pause, resume (even from another device), cancel, retry failed, and export CSV.
  BCC broadcasts are split into groups of 90, because Gmail rejects more than 100 recipients per message.
- **Sent history and audit log.** Every sign-in, account change, template share and send is recorded.

## Stack

Vite + React + TypeScript (Tailwind, shadcn/ui) · Hono API bundled into a single Vercel function ·
Neon Postgres via Drizzle · Nodemailer over `smtp.gmail.com:465` · TipTap editor · juice and html-to-text.

```
src/      React app (pages, compose/, editor/, preview/, templates/, admin/, history/)
server/   Hono app, routes/, lib/ (render pipeline, jobs, mailer, lint, uploads, auth), db/ (schema + migrations),
          dev.ts (local API server), vercel.ts (production function entry)
shared/   Code used by both sides: API types, themes, fonts, recipients/CC-BCC rules, text helpers
scripts/  create-admin, migrate, build-vercel (production build), build-caniemail
tests/    Vitest suites
```

## Setup

1. `npm install`
2. Copy `.env.example` to `.env` and fill it in. **Never change `APP_ENCRYPTION_KEY` once sender accounts exist.**
3. `npm run db:migrate` creates the tables.
4. `npm run create-admin -- --email you@kiit.ac.in --name "Your Name"` creates the first administrator. You
   are prompted for a password, or add `--generate` for a temporary one that must be changed at first sign-in.
5. `npm run dev` starts the app at http://localhost:5173 (the API runs on :8787 and is proxied).
6. Sign in, open **Sender accounts**, add an `@kiit.ac.in` account with its App Password, then add users and
   allot accounts to them.

**Creating a Google App Password:** sign in to the sender mailbox, then go to Google Account → Security →
2-Step Verification (it must be on) → App passwords (myaccount.google.com/apppasswords) and create one named
"KIIT Mailer". Paste the 16 letters into the app.

| Command | What it does |
|---|---|
| `npm run dev` | API + web dev servers |
| `npm test` | All unit tests |
| `npm run typecheck` | TypeScript check of the whole repo |
| `npm run build` | Production build of the web app (to `dist/`) |
| `npm run db:generate` / `db:migrate` | Create / apply a migration after editing `server/db/schema.ts` |

## Deploying on Vercel

`npm run build` produces a complete [Build Output API](https://vercel.com/docs/build-output-api/v3) folder
(`.vercel/output`): the static site plus **one** Node.js function containing the whole API, already bundled to
plain JavaScript (no native modules), with routing and security headers (CSP, HSTS, frame denial). Vercel deploys
it as-is; `vercel.json` only sets the build command.

Environment variables (Project → Settings → Environment Variables):

| Variable | Value |
|---|---|
| `DATABASE_URL` | Neon **production** branch (`main`), pooled connection string |
| `APP_ENCRYPTION_KEY` | 32 random bytes, base64. Use a different key from dev, and never change it once sender accounts exist |
| `APP_BASE_URL` | The public URL, e.g. `https://mailer.example.com` (falls back to the Vercel production domain) |
| `BLOB_READ_WRITE_TOKEN` | A **private** Vercel Blob store for attachments |
| `BLOB_PUBLIC_READ_WRITE_TOKEN` | Optional. A second, **public** Blob store for images inside email bodies; without it, body images are embedded as inline attachments |

Before the first deploy, run `npm run db:migrate` and `npm run create-admin` against the production database
(set `DATABASE_URL` for that shell).

Sending runs while the sender's browser tab is open: the page calls the API in short batches (about 30 seconds
each, well inside the function time limit). Closing the tab pauses the job, and it can be resumed from Sent history.

## Deliverability notes for KIIT IT

kiit.ac.in already publishes SPF (`include:_spf.google.com`) and a Google DKIM key, so mail sent through Gmail
SMTP is authenticated. DMARC is currently `p=none; pct=1`. After monitoring the reports, moving to
`p=quarantine` (with `rua` reporting) would further protect the domain and improve inbox placement. Gmail limits
each Workspace account to 2,000 recipients per day; the app enforces each account's limit.
