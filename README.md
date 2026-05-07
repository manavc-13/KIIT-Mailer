# KIIT Mailer

A lightweight, local web-based mailing utility for KIIT University. Compose HTML
emails (with plain-text fallback), send to a single recipient or in bulk via
CSV upload or an in-browser editable data grid, and preview exactly how the
email will look in clients like Gmail — across light/dark themes and
desktop/mobile viewports.

## Features

- **HTML-First Composer** — primary editor for raw HTML, with a Rich Text mode
  (Quill) and a secondary Plain-Text editor for the text MIME part.
- **Single & Bulk Send** —
  - **Single Recipient**: send to one address with optional custom placeholder
    fields you define inline.
  - **Bulk via CSV**: upload a CSV with `Name`, `Email`, and any extra columns.
  - **Bulk via Manual Entry**: an editable spreadsheet-like grid in the
    browser with add/remove rows, custom columns (= custom placeholders),
    and inline validation.
- **Placeholders** — any column becomes a `{Column}` placeholder usable in
  Subject, HTML, and Plain Text.
- **Email-Client-Accurate Preview** — Gmail-style chrome (sender, subject,
  recipient, date) wrapping a sandboxed iframe of the actual HTML.
  - **Light / Dark theme** toggle (simulated)
  - **Desktop / Mobile** viewport toggle (mobile shows a phone frame at 380px)
  - **Per-recipient preview** when a list is loaded
- **Attachments** — drag/drop multiple files, total size guard (25 MB).
- **Activity Log** — local timestamped log of sends and errors.
- **Local Settings** — SMTP credentials (KIIT email + Google App Password) are
  stored only in your browser's `localStorage` and sent directly to the local
  server when sending.

## Quick Start

```bash
npm install
npm start
```

Then open the URL printed in the console (typically `http://localhost:3000`).

1. On first launch, open **Settings** and enter:
   - your `*@kiit.ac.in` email
   - the Google App Password (see welcome page for steps)
   - optional display name and reply-to
2. Switch to **Compose**, pick **Single Recipient** or **Bulk Send**.
3. Compose the body in HTML (default), Rich Text, or Plain Text.
4. Click **👁️ Preview** to see a Gmail-style preview with theme/viewport
   toggles. Use the recipient selector to preview as any row.
5. Click **Send Email** / **Send Bulk Emails**.

## Notes

- Credentials never leave your machine other than the SMTP send itself.
- The HTML body is sent as `html`; the Plain-Text body, when present, is sent
  as the `text` MIME part for clients that prefer it.
- Built with Node/Express, Nodemailer, Busboy, PapaParse, and Quill.
