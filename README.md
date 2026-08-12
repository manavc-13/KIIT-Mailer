# KIIT Mailer

A web-based mailing utility for KIIT University IQAC. Compose HTML emails
(with plain-text fallback), send to a single recipient or in bulk via CSV
upload or an in-browser editable data grid — with **per-recipient CC/BCC** —
and preview exactly how the email will look in clients like Gmail before you
send.

Runs locally (`npm start` / the packaged `.exe`) or deployed on Vercel.

## Features

- **3-step compose flow** — Recipients → Compose → Review & Send, with a
  persistent live preview pane throughout.
- **HTML-First Composer** — primary editor for raw HTML, with a Rich Text mode
  (Quill) and a secondary Plain-Text editor for the text MIME part.
- **Single & Bulk Send** —
  - **Single Recipient**: send to one address with optional custom placeholder
    fields you define inline.
  - **Bulk via CSV**: upload a CSV. Columns are auto-detected case-insensitively
    (`Email`, `E-mail`, `email address`, … all work) with an editable mapping
    panel if auto-detection needs correcting.
  - **Bulk via Manual Entry**: an editable spreadsheet-like grid in the
    browser with add/remove rows, custom columns, and inline validation.
- **Per-recipient CC/BCC** — add `CC` and `BCC` columns to your CSV or grid;
  each row's copies are merged with the global CC/BCC fields, de-duplicated,
  and never repeat an address already in To. See [CC/BCC columns](#ccbcc-columns)
  below.
- **Placeholders** — any column becomes a `{Column}` placeholder usable in
  Subject, HTML, and Plain Text, resolved case-insensitively.
- **Templates & drafts** — save/load/export/import named templates; your
  in-progress compose is autosaved and restored after an accidental refresh.
- **Email-Client-Accurate Preview** — Gmail-style chrome (sender, subject,
  recipient, date) wrapping a sandboxed iframe of the actual HTML, updated
  live as you type.
  - **Light / Forced-Dark** toggle (simulates a mail client re-theming your
    email; the email itself always renders light — see [Notes](#notes))
  - **Desktop / Mobile** viewport toggle
  - **Per-recipient navigation** (◀ / ▶) when a list is loaded
- **Attachments sized against Gmail's real limit** — a live meter shows raw
  size, base64-encoded size, and % of Gmail's 25 MB message cap, with a soft
  warning above 10 MB. See [Attachment limits](#attachment-limits).
- **Batched, resumable bulk sends** — up to 10 recipients per request over a
  pooled SMTP connection (instead of one connection per email), automatic
  retry with backoff on transient Gmail throttling, pause/cancel/resume
  across a browser refresh, and a **Retry failed only** button.
- **Plain-English send errors** — bad App Password, daily quota exhausted,
  message too large, etc. are recognized and explained instead of showing raw
  SMTP text.
- **Activity Log** — local timestamped log of sends and errors.
- **Local Settings** — SMTP credentials (KIIT email + Google App Password) are
  stored only in your browser's `localStorage` and sent directly to the mail
  server when sending.

## Quick Start

```bash
npm install
npm start
```

Then open the URL printed in the console (typically `http://localhost:3000`).

1. On first launch, the **Settings** drawer opens automatically. Enter:
   - your `*@kiit.ac.in` email
   - the Google App Password (see the welcome page for steps)
   - optional display name and reply-to
2. **Recipients** step — pick Single / CSV / Manual Grid, and (for CSV/Grid)
   confirm the column mapping and global CC/BCC.
3. **Compose** step — write the subject/body, attach files, optionally save
   as a template.
4. **Review & Send** step — check the pre-flight summary and quota, then
   **Send test to me** or **Send**.

## CC/BCC columns

Add a `CC` and/or `BCC` column to your CSV or manual grid (the manual grid
creates them by default). Each recipient's row can carry its own copy
addresses — separate multiple addresses in one cell with `;` (or a comma, if
you quote the cell). At send time, for every recipient:

```
final CC  = row's CC ∪ global CC field   (de-duplicated, minus the To address)
final BCC = row's BCC ∪ global BCC field (de-duplicated, minus To and final CC)
```

The global CC/BCC fields (Recipients step) also support placeholders, e.g.
`{Supervisor}`, if that column exists.

Column names don't have to be exactly `CC`/`BCC` — the mapping panel
auto-detects common variants (`cc email`, `copy`, `bcc address`, …) and lets
you remap manually.

## Attachment limits

Gmail (and Google Workspace) caps the total **encoded** message size at
**25 MB**; base64 inflates raw attachment bytes by ~4/3, so the real usable
raw-attachment budget is roughly **18 MB**. The app enforces this — not an
arbitrary flat cap — and shows both numbers live as you attach files, with a
soft warning above 10 MB (many external mail servers reject messages smaller
than Gmail's own limit).

Attachments are uploaded **once** (not once per recipient) and referenced by
every message in a bulk send, which is also what fixes the old "payload
error under 15 MB" bug: that was Vercel's serverless function body cap
(4.5 MB) being hit because attachments used to travel inside every single
send request. Now:

- **Deployed on Vercel with a Blob store configured** (`BLOB_READ_WRITE_TOKEN`
  set — create one under Storage → Blob in the Vercel dashboard): attachments
  upload directly from the browser to Blob storage, up to the full ~18 MB
  budget, and are deleted after you remove them or close the tab.
- **Running locally / the packaged `.exe`**: attachments upload to a local
  temp directory; same ~18 MB budget.
- **Deployed on Vercel without a Blob store**: a degraded fallback embeds
  small attachments (~2.8 MB raw) directly in the send request. The app shows
  a banner explaining this and how to remove the limit (configure Blob
  storage).

## Notes

- Credentials never leave your machine other than the SMTP send itself (and,
  if deployed, the request to your own Vercel function).
- The HTML body is sent as `html`; the Plain-Text body, when present, is sent
  as the `text` MIME part for clients that prefer it.
- Outgoing HTML always opts out of mail-client dark-mode re-theming
  (`color-scheme: light only`), so your design renders exactly as composed
  regardless of the recipient's theme. The preview's "Forced Dark" toggle
  simulates what an email *without* that opt-out would look like — for
  reference only, it does not change what's sent.
- Built with Node/Express, Nodemailer, PapaParse, and Quill on the frontend;
  Vercel Blob (optional) for attachment storage.
