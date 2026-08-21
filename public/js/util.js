// Pure helper functions — no DOM, no state. Carried over from the original
// app.js (behaviour preserved) with two additions: substitute() now resolves
// placeholders case-insensitively, and a couple of small formatting helpers.

export function escapeAttr(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

export function parseEmailList(value) {
    return String(value || '')
        .split(/[\s,;]+/)
        .map(v => v.trim())
        .filter(Boolean);
}

export function normalizeEmailList(value) {
    return parseEmailList(value).join(', ');
}

export function invalidEmails(value) {
    return parseEmailList(value).filter(email => !/^\S+@\S+\.\S+$/.test(email));
}

export function hasValidEmailList(value) {
    const emails = parseEmailList(value);
    return emails.length > 0 && invalidEmails(value).length === 0;
}

// Case-insensitive, de-duplicated union of one or more email-list strings.
export function uniqueEmails(...lists) {
    const seen = new Set();
    const out = [];
    lists.forEach(list => {
        parseEmailList(list).forEach(e => {
            const key = e.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            out.push(e);
        });
    });
    return out;
}

// Placeholder substitution: {Key} or {key} both resolve against a row whose
// keys may be any case. Exact-case key wins if present; otherwise the first
// case-insensitive match is used. Unresolved tokens are left verbatim.
export function substitute(str, row) {
    if (!str) return str;
    const lowerIndex = {};
    Object.keys(row || {}).forEach(k => {
        const lk = k.trim().toLowerCase();
        if (!(lk in lowerIndex)) lowerIndex[lk] = row[k];
    });
    return String(str).replace(/\{([^{}]+)\}/g, (_m, key) => {
        const trimmed = key.trim();
        if (row && row[trimmed] != null) return row[trimmed];
        const lower = trimmed.toLowerCase();
        if (lowerIndex[lower] != null) return lowerIndex[lower];
        return `{${key}}`;
    });
}

export function textToHtml(text) {
    const escaped = String(text || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
    const body = escaped.split(/\n/).map(l => l || '&nbsp;').join('<br>');
    return `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light"></head>
<body style="margin:0; padding:24px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif; font-size:14px; line-height:1.6; color:#1f2328; background:#ffffff;">
<div style="max-width:600px; margin:0;">${body}</div>
</body></html>`;
}

export function wrapRichContent(html) {
    if (/<!DOCTYPE/i.test(html)) return html;
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light only">
<meta name="supported-color-schemes" content="light">
<style>
  :root { color-scheme: light only; }
  body { margin:0; padding:24px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif; color:#1f2328; background:#ffffff; }
</style>
</head>
<body>
<div style="max-width:600px; margin:0;">${html}</div>
</body>
</html>`;
}

// Ensure outgoing HTML opts out of mail-client dark-mode transformations.
// Idempotent: safe to call on HTML that already declares color-scheme.
export function lockLightMode(html) {
    const lockMetas = [
        '<meta name="color-scheme" content="light only">',
        '<meta name="supported-color-schemes" content="light">',
    ];
    const lockStyle = '<style>:root{color-scheme:light only;supported-color-schemes:light}body{color-scheme:light only}</style>';

    const hasColorScheme = /<meta[^>]+name=["']color-scheme["']/i.test(html);
    const hasSupported = /<meta[^>]+name=["']supported-color-schemes["']/i.test(html);
    const hasLockStyle = /color-scheme\s*:\s*light\s*only/i.test(html);

    const inject =
        (hasColorScheme ? '' : lockMetas[0]) +
        (hasSupported ? '' : lockMetas[1]) +
        (hasLockStyle ? '' : lockStyle);

    if (!inject) return html;

    if (/<head[^>]*>/i.test(html)) {
        return html.replace(/<head[^>]*>/i, m => m + inject);
    }
    if (/<html[^>]*>/i.test(html)) {
        return html.replace(/<html[^>]*>/i, m => m + '<head>' + inject + '</head>');
    }
    return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">${inject}</head><body>${html}</body></html>`;
}

export function injectPreheader(html, preheader) {
    const safe = preheader
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const block = `<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;font-size:1px;line-height:1px;color:transparent;opacity:0;">${safe}</div>`;
    if (/<body[^>]*>/i.test(html)) {
        return html.replace(/<body[^>]*>/i, m => m + block);
    }
    return block + html;
}

export function dedupeRecipients(list) {
    const seen = new Set();
    const out = [];
    list.forEach(r => {
        const key = (r.Email || '').trim().toLowerCase();
        if (!key) return;
        if (seen.has(key)) return;
        seen.add(key);
        out.push(r);
    });
    return out;
}

// Tabular paste parser: detects tab-separated or comma-separated.
export function parseTabular(text) {
    const lines = text.replace(/\r/g, '').split('\n').filter(l => l.length > 0);
    if (lines.length === 0) return { columns: [], rows: [] };
    const sep = lines[0].includes('\t') ? '\t' : ',';
    const splitLine = (line) => {
        if (sep === '\t') return line.split('\t').map(s => s.trim());
        const result = []; let cur = ''; let inQ = false;
        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (inQ) {
                if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
                else if (ch === '"') { inQ = false; }
                else cur += ch;
            } else {
                if (ch === '"') inQ = true;
                else if (ch === ',') { result.push(cur.trim()); cur = ''; }
                else cur += ch;
            }
        }
        result.push(cur.trim());
        return result;
    };
    const headers = splitLine(lines[0]);
    const rows = lines.slice(1).map(line => {
        const cells = splitLine(line);
        const o = {};
        headers.forEach((h, i) => { o[h] = cells[i] != null ? cells[i] : ''; });
        return o;
    });
    return { columns: headers, rows };
}

export function toCSV(columns, rows) {
    const esc = v => {
        const s = String(v == null ? '' : v);
        if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
        return s;
    };
    const head = columns.map(esc).join(',');
    const body = rows.map(r => columns.map(c => esc(r[c])).join(',')).join('\n');
    return head + '\n' + body + '\n';
}

export function downloadFile(content, filename, mime) {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
}

export function formatBytes(bytes) {
    return (bytes / 1024 / 1024).toFixed(2) + ' MB';
}

export function encodedSize(rawBytes) {
    return Math.ceil(rawBytes / 3) * 4;
}
