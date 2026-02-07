# IQAC Mailer (Web Version)

A modern, web-based bulk email system for IQAC, replacing the legacy Windows application. 
Built with Node.js (Vercel Functions) and Vanilla HTML/CSS/JS.

## Features

- **Secure Login**: Staff ID-based authentication (hashed validation).
- **Modern UI**: Sleek dark theme using Inter font.
- **Bulk Mailing**: Upload CSV, map placeholders (e.g., `{Name}`), and send thousands of emails.
- **Rich Text**: Basic formatting (Bold, Italic, Link) for email bodies.
- **Attachments**: Support for multiple file attachments.
- **Vercel Ready**: Structure optimized for zero-config Vercel deployment.

## Local Development

1.  **Install Dependencies**:
    ```bash
    npm install
    ```

2.  **Configure Environment**:
    -   Rename `env.example` to `.env` (or create one).
    -   Add your SMTP details:
        ```ini
        SMTP_HOST=smtp.gmail.com
        SMTP_PORT=465
        SMTP_USER=your-email@gmail.com
        SMTP_PASS=your-app-password
        ALLOWED_HASHES=5994471abb01112afcc18159f6cc74b4f511b99806da59b3caf5a9c173cacfc5
        ```
    -   *Note: The default hash is for Staff ID "12345".*

3.  Run `node local_server.js` (Recommended if Vercel CLI fails) or `vercel dev`.
4.  Open `http://localhost:3000`.

## Deployment (Vercel)

1.  Push this folder to GitHub.
2.  Import the project into Vercel.
3.  Vercel will automatically detect the configuration.
4.  **Important**: Add the Environment Variables (`SMTP_HOST`, `SMTP_PASS`, etc.) in the Vercel Project Settings.
5.  Deploy!

## CSV Format

For bulk emails, upload a CSV with at least an `Email` column.

```csv
Name,Email,Department
John Doe,john@example.com,CSE
Jane Smith,jane@example.com,ECE
```

You can use placeholders like `{Name}` or `{Department}` in your email body.
