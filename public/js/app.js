// KIIT Mailer - Frontend Logic
// HTML-first email composer with Single & Bulk (CSV / Manual Grid) modes,
// Gmail-style preview with theme + viewport toggles.

const MAX_SIZE_MB = 25;
const STORAGE_SETTINGS = 'kiit_mailer_settings';
const STORAGE_LOGS = 'kiit_mailer_logs';
const STORAGE_QUEUE = 'kiit_mailer_pending_queue';

// Common spam-trigger words (subject) — not exhaustive, just a friendly nudge.
const SPAM_WORDS = [
    'free', 'winner', 'won', 'prize', 'urgent', 'cash', 'click here', 'buy now',
    'limited time', 'act now', 'guarantee', 'risk-free', 'congratulations',
    'lottery', 'discount', 'offer', '!!', '$$', '100%'
];

// ---------- Default Starter HTML (generic, KIIT-friendly, light-locked) ----------
// IMPORTANT: This template intentionally opts OUT of automatic dark-mode color
// transformation by mail clients. Your colors will appear exactly as designed.
const DEFAULT_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light only">
<meta name="supported-color-schemes" content="light">
<title>KIIT Mailer</title>
<style>
  :root { color-scheme: light only; }
  body { margin:0; padding:0; -webkit-text-size-adjust:100%; }
  @media only screen and (max-width:600px) {
    .card { width:100% !important; padding:24px !important; border-radius:0 !important; }
    .btn { width:100% !important; box-sizing:border-box !important; }
  }
</style>
</head>
<body class="body-bg" style="margin:0; padding:0; background-color:#f4f5f7;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="body-bg" style="background-color:#f4f5f7;">
    <tr>
      <td align="center" style="padding:32px 16px;">
        <table role="presentation" class="card" width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px; width:100%; background-color:#ffffff; border-radius:12px; overflow:hidden; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color:#1f2328;">
          <tr>
            <td style="padding:32px 36px 8px 36px;">
              <h1 class="accent" style="margin:0 0 6px 0; font-size:22px; font-weight:700; color:#0b5394; letter-spacing:-0.2px;">
                KIIT University
              </h1>
              <p class="muted" style="margin:0; font-size:14px; color:#6b7280;">
                Bhubaneswar, Odisha &middot; kiit.ac.in
              </p>
            </td>
          </tr>
          <tr><td style="padding:16px 36px 0 36px;"><hr class="divider" style="border:none; border-top:1px solid #e5e7eb; margin:0;"></td></tr>
          <tr>
            <td style="padding:24px 36px;">
              <p style="margin:0 0 14px 0; font-size:16px; line-height:1.6;">Dear {Name},</p>
              <p style="margin:0 0 14px 0; font-size:16px; line-height:1.6;">
                Replace this with your message. You can use placeholders such as
                <strong>{Name}</strong> and <strong>{Email}</strong> — they will be
                substituted per recipient when sending.
              </p>
              <p style="margin:0 0 24px 0; font-size:16px; line-height:1.6;">
                Best regards,<br>
                <strong>KIIT University</strong>
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td bgcolor="#0b5394" class="btn" style="border-radius:6px;">
                    <a href="https://kiit.ac.in" target="_blank" style="display:inline-block; padding:12px 22px; font-size:14px; font-weight:600; color:#ffffff; text-decoration:none;">Visit KIIT</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td class="muted" style="padding:16px 36px 28px 36px; font-size:12px; color:#6b7280; border-top:1px solid #e5e7eb;">
              You are receiving this email from KIIT Mailer.
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

const DEFAULT_TEXT = `Dear {Name},

Replace this with your plain-text message. Placeholders like {Name} and {Email} will be substituted per recipient.

Best regards,
KIIT University`;

// ---------- DOM Elements ----------
const els = {
    tabs: document.querySelectorAll('.nav-btn'),
    contents: document.querySelectorAll('.tab-content'),

    // Send Mode
    sendModeSeg: document.getElementById('sendModeSeg'),
    singleFields: document.getElementById('singleFields'),
    bulkFields: document.getElementById('bulkFields'),
    singleEmail: document.getElementById('singleEmail'),
    singleName: document.getElementById('singleName'),
    singleCustomFields: document.getElementById('singleCustomFields'),
    addSingleFieldBtn: document.getElementById('addSingleFieldBtn'),

    // Bulk Source
    bulkSourceSeg: document.getElementById('bulkSourceSeg'),
    bulkCsv: document.getElementById('bulkCsv'),
    bulkManual: document.getElementById('bulkManual'),

    // CSV
    csvFile: document.getElementById('csvFile'),
    csvStatus: document.getElementById('csvStatus'),
    downloadTmplBtn: document.getElementById('downloadTemplateBtn'),

    // Manual grid
    manualGrid: document.getElementById('manualGrid'),
    addRowBtn: document.getElementById('addRowBtn'),
    addColBtn: document.getElementById('addColBtn'),
    clearGridBtn: document.getElementById('clearGridBtn'),
    gridStatus: document.getElementById('gridStatus'),
    deleteSelectedBtn: document.getElementById('deleteSelectedBtn'),
    selectedCount: document.getElementById('selectedCount'),
    pasteImportBtn: document.getElementById('pasteImportBtn'),
    pasteImportArea: document.getElementById('pasteImportArea'),
    pasteImportText: document.getElementById('pasteImportText'),
    pasteImportApply: document.getElementById('pasteImportApply'),
    pasteImportCancel: document.getElementById('pasteImportCancel'),
    pasteImportAppend: document.getElementById('pasteImportAppend'),
    exportGridBtn: document.getElementById('exportGridBtn'),

    // Common
    subject: document.getElementById('subject'),
    subjectCounter: document.getElementById('subjectCounter'),
    subjectWarn: document.getElementById('subjectWarn'),
    preheader: document.getElementById('preheader'),
    preheaderCounter: document.getElementById('preheaderCounter'),
    placeholderToolbar: document.getElementById('placeholderToolbar'),
    attachInput: document.getElementById('attachments'),
    attachList: document.getElementById('attachmentList'),
    clearAttachmentsBtn: document.getElementById('clearAttachmentsBtn'),
    sendBtn: document.getElementById('sendBtn'),
    sendTestBtn: document.getElementById('sendTestBtn'),

    // Bulk options
    bulkOptions: document.getElementById('bulkOptions'),
    delayMs: document.getElementById('delayMs'),
    dedupToggle: document.getElementById('dedupToggle'),
    runPreflightBtn: document.getElementById('runPreflightBtn'),
    refreshPreflightBtn: document.getElementById('refreshPreflightBtn'),
    preflightPanel: document.getElementById('preflightPanel'),
    preflightBody: document.getElementById('preflightBody'),

    // Resume
    resumeBanner: document.getElementById('resumeBanner'),
    resumeMeta: document.getElementById('resumeMeta'),
    resumeBtn: document.getElementById('resumeBtn'),
    discardResumeBtn: document.getElementById('discardResumeBtn'),

    // Send console
    pauseBtn: document.getElementById('pauseBtn'),
    resumeBtnConsole: document.getElementById('resumeBtnConsole'),
    cancelBtn: document.getElementById('cancelBtn'),
    successCount: document.getElementById('successCount'),
    failureCount: document.getElementById('failureCount'),
    pendingCount: document.getElementById('pendingCount'),
    resultsTable: document.getElementById('resultsTable'),
    exportResultsBtn: document.getElementById('exportResultsBtn'),
    closeConsoleBtn: document.getElementById('closeConsoleBtn'),
    consoleTitle: document.getElementById('consoleTitle'),

    // Editor
    editorSeg: document.getElementById('editorSeg'),
    htmlEditor: document.getElementById('htmlEditor'),
    textEditor: document.getElementById('textEditor'),
    richEditorContainer: document.getElementById('richEditorParams'),
    htmlEditorContainer: document.getElementById('htmlEditorParams'),
    textEditorContainer: document.getElementById('textEditorParams'),

    // Preview
    previewBtn: document.getElementById('previewBtn'),
    previewModal: document.getElementById('previewModal'),
    closePreviewBtn: document.getElementById('closePreviewBtn'),
    previewFrame: document.getElementById('previewFrame'),
    previewStage: document.getElementById('previewStage'),
    previewThemeSeg: document.getElementById('previewThemeSeg'),
    previewViewportSeg: document.getElementById('previewViewportSeg'),
    previewRecipient: document.getElementById('previewRecipient'),
    gmailFrame: document.getElementById('gmailFrame'),
    gmSubject: document.getElementById('gmSubject'),
    gmFromName: document.getElementById('gmFromName'),
    gmFromEmail: document.getElementById('gmFromEmail'),
    gmTo: document.getElementById('gmTo'),
    gmAvatar: document.getElementById('gmAvatar'),
    gmDate: document.getElementById('gmDate'),
    // Warnings
    modeWarningModal: document.getElementById('modeWarningModal'),
    confirmModeSwitch: document.getElementById('confirmModeSwitch'),
    cancelModeSwitch: document.getElementById('cancelModeSwitch'),

    // Progress
    overlay: document.getElementById('progressOverlay'),
    progressBar: document.getElementById('progressBar'),
    progressText: document.getElementById('progressText'),
    progressPercent: document.getElementById('progressPercent'),

    // Logs
    logTerminal: document.getElementById('logTerminal'),
    clearLogsBtn: document.getElementById('clearLogs'),

    // Toasts
    toastContainer: document.getElementById('toastContainer'),

    // User
    userNameDisplay: document.getElementById('userNameDisplay'),
    userAvatar: document.getElementById('userAvatar'),

    // Settings
    settingsEmail: document.getElementById('settingsEmail'),
    settingsPass: document.getElementById('settingsPass'),
    settingsDisplayName: document.getElementById('settingsDisplayName'),
    settingsReplyTo: document.getElementById('settingsReplyTo'),
    saveSettingsBtn: document.getElementById('saveSettingsBtn'),
    settingsTabBtn: document.getElementById('settingsTabBtn'),
};

// ---------- State ----------
const state = {
    sendMode: 'single',     // 'single' | 'bulk'
    bulkSource: 'csv',      // 'csv' | 'manual'
    editorMode: 'html',     // 'html' | 'rich' | 'text'
    pendingEditorMode: null,

    // Single recipient extra fields (array of {key, value})
    singleExtras: [],

    // Manual grid: columns + rows
    manualColumns: ['Name', 'Email'],
    manualRows: [{ Name: '', Email: '' }],

    // CSV
    csvData: null,
    csvHeaders: [],

    // Attachments
    attachments: [],

    // Misc
    isSending: false,
    isPaused: false,
    isCancelled: false,
    quill: null,
    credentials: null,

    // Manual grid selection
    selectedRows: new Set(),

    // Send job
    sendResults: [],   // [{index, email, status, detail}]
    sendQueue: [],     // remaining valid recipients (with subject/html/text resolved)
    sendCursor: 0,

    // Preview
    previewTheme: 'light',
    previewViewport: 'desktop',
    previewRowIndex: 0,
};

// ---------- Init ----------
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('year').textContent = new Date().getFullYear();

    loadSettings();
    initQuill();

    setupTabs();
    setupSendModeToggle();
    setupBulkSourceToggle();
    setupSingleExtras();
    setupCSV();
    setupManualGrid();
    setupEditorMode();
    setupAttachments();
    setupSending();
    setupLogs();
    setupSettings();
    setupPreview();
    setupSubjectHelpers();
    setupBulkOptions();
    setupSendConsole();
    setupTestSend();
    checkPendingQueue();

    renderManualGrid();
    refreshPlaceholders();
    loadHistory();
    updateSubjectHelpers();
    updatePreheaderCounter();

    if (!state.credentials || !state.credentials.email || !state.credentials.pass) {
        showToast('Please configure settings first', 'warning');
        setTimeout(() => els.settingsTabBtn.click(), 400);
    }
});

// ---------- Settings ----------
function loadSettings() {
    const stored = localStorage.getItem(STORAGE_SETTINGS);
    if (!stored) return;
    try {
        state.credentials = JSON.parse(stored);
        els.settingsEmail.value = state.credentials.email || '';
        els.settingsPass.value = state.credentials.pass || '';
        els.settingsDisplayName.value = state.credentials.displayName || '';
        els.settingsReplyTo.value = state.credentials.replyTo || '';
        if (state.credentials.email) {
            const dn = state.credentials.displayName || state.credentials.email.split('@')[0];
            els.userNameDisplay.textContent = dn;
            els.userAvatar.textContent = dn.charAt(0).toUpperCase();
        }
    } catch (e) {
        console.error('Failed to parse settings', e);
    }
}

function setupSettings() {
    els.saveSettingsBtn.addEventListener('click', () => {
        const email = els.settingsEmail.value.trim();
        const pass = els.settingsPass.value.trim();
        const displayName = els.settingsDisplayName.value.trim();
        const replyTo = els.settingsReplyTo.value.trim();

        if (!email.endsWith('@kiit.ac.in')) {
            showToast('Email must be @kiit.ac.in', 'error');
            return;
        }
        if (!pass) {
            showToast('App Password is required', 'error');
            return;
        }

        const creds = { email, pass, displayName, replyTo };
        localStorage.setItem(STORAGE_SETTINGS, JSON.stringify(creds));
        state.credentials = creds;

        const dn = displayName || email.split('@')[0];
        els.userNameDisplay.textContent = dn;
        els.userAvatar.textContent = dn.charAt(0).toUpperCase();

        showToast('Settings saved', 'success');
    });
}

// ---------- Tabs ----------
function setupTabs() {
    els.tabs.forEach(btn => {
        btn.addEventListener('click', () => {
            els.tabs.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const target = btn.dataset.tab;
            els.contents.forEach(c => c.classList.remove('active'));
            document.getElementById(target).classList.add('active');
        });
    });
}

// ---------- Send Mode Toggle ----------
function setupSendModeToggle() {
    els.sendModeSeg.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            els.sendModeSeg.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.sendMode = btn.dataset.mode;
            els.singleFields.style.display = state.sendMode === 'single' ? 'block' : 'none';
            els.bulkFields.style.display = state.sendMode === 'bulk' ? 'block' : 'none';
            els.bulkOptions.style.display = state.sendMode === 'bulk' ? 'flex' : 'none';
            els.sendBtn.textContent = state.sendMode === 'single' ? 'Send Email' : 'Send Bulk Emails';
            refreshPlaceholders();
        });
    });
}

// ---------- Single Extras ----------
function setupSingleExtras() {
    els.addSingleFieldBtn.addEventListener('click', () => {
        state.singleExtras.push({ key: '', value: '' });
        renderSingleExtras();
    });
}

function renderSingleExtras() {
    els.singleCustomFields.innerHTML = '';
    state.singleExtras.forEach((f, idx) => {
        const row = document.createElement('div');
        row.className = 'custom-field-row';
        row.innerHTML = `
            <input type="text" placeholder="Field name (e.g. Department)" value="${escapeAttr(f.key)}" data-i="${idx}" data-k="key">
            <input type="text" placeholder="Value" value="${escapeAttr(f.value)}" data-i="${idx}" data-k="value">
            <button type="button" class="icon-btn" data-remove="${idx}" title="Remove">✕</button>
        `;
        els.singleCustomFields.appendChild(row);
    });
    els.singleCustomFields.querySelectorAll('input').forEach(inp => {
        inp.addEventListener('input', e => {
            const i = +e.target.dataset.i;
            const k = e.target.dataset.k;
            state.singleExtras[i][k] = e.target.value;
            refreshPlaceholders();
        });
    });
    els.singleCustomFields.querySelectorAll('[data-remove]').forEach(btn => {
        btn.addEventListener('click', e => {
            const i = +e.currentTarget.dataset.remove;
            state.singleExtras.splice(i, 1);
            renderSingleExtras();
            refreshPlaceholders();
        });
    });
}

function escapeAttr(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

// ---------- Bulk Source Toggle ----------
function setupBulkSourceToggle() {
    els.bulkSourceSeg.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            els.bulkSourceSeg.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.bulkSource = btn.dataset.source;
            els.bulkCsv.style.display = state.bulkSource === 'csv' ? 'block' : 'none';
            els.bulkManual.style.display = state.bulkSource === 'manual' ? 'block' : 'none';
            refreshPlaceholders();
        });
    });
}

// ---------- CSV ----------
function setupCSV() {
    els.csvFile.addEventListener('change', e => {
        const file = e.target.files[0];
        if (file) parseCSV(file);
    });

    els.downloadTmplBtn.addEventListener('click', () => {
        const csv = 'Name,Email\nJohn Doe,john.doe@example.com\nJane Roe,jane.roe@example.com\n';
        const blob = new Blob([csv], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'kiit_mailer_template.csv'; a.click();
        URL.revokeObjectURL(url);
    });
}

function parseCSV(file) {
    els.csvStatus.textContent = 'Parsing...';
    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: results => {
            if (results.data && results.data.length > 0) {
                state.csvData = results.data;
                state.csvHeaders = results.meta.fields;
                els.csvStatus.textContent = `Loaded ${results.data.length} recipients · Columns: ${state.csvHeaders.join(', ')}`;
                refreshPlaceholders();
                showToast(`CSV loaded: ${results.data.length} rows`, 'success');
            } else {
                els.csvStatus.textContent = 'Error: No data found';
                showToast('CSV parse error: empty or invalid file', 'error');
            }
        },
        error: err => {
            els.csvStatus.textContent = 'Error';
            showToast(`CSV error: ${err.message}`, 'error');
        }
    });
}

// ---------- Manual Grid ----------
function setupManualGrid() {
    els.addRowBtn.addEventListener('click', () => {
        const row = {};
        state.manualColumns.forEach(c => row[c] = '');
        state.manualRows.push(row);
        renderManualGrid();
    });

    els.addColBtn.addEventListener('click', () => {
        const name = prompt('New column name (becomes a placeholder like {Name}):');
        if (!name) return;
        const clean = name.trim();
        if (!clean) return;
        if (state.manualColumns.includes(clean)) {
            showToast('Column already exists', 'warning');
            return;
        }
        state.manualColumns.push(clean);
        state.manualRows.forEach(r => r[clean] = '');
        renderManualGrid();
        refreshPlaceholders();
    });

    els.clearGridBtn.addEventListener('click', () => {
        if (!confirm('Clear all manual entries?')) return;
        state.manualColumns = ['Name', 'Email'];
        state.manualRows = [{ Name: '', Email: '' }];
        state.selectedRows.clear();
        renderManualGrid();
        refreshPlaceholders();
    });

    // Delete selected (multi-select)
    els.deleteSelectedBtn.addEventListener('click', () => {
        const sel = [...state.selectedRows].sort((a, b) => b - a);
        if (sel.length === 0) return;
        if (!confirm(`Delete ${sel.length} selected row(s)?`)) return;
        sel.forEach(i => state.manualRows.splice(i, 1));
        state.selectedRows.clear();
        if (state.manualRows.length === 0) {
            const empty = {}; state.manualColumns.forEach(c => empty[c] = '');
            state.manualRows.push(empty);
        }
        renderManualGrid();
        showToast(`Deleted ${sel.length} row(s)`, 'success');
    });

    // Paste-from-Excel modal area
    els.pasteImportBtn.addEventListener('click', () => {
        const open = els.pasteImportArea.style.display !== 'none';
        els.pasteImportArea.style.display = open ? 'none' : 'block';
        if (!open) setTimeout(() => els.pasteImportText.focus(), 50);
    });
    els.pasteImportCancel.addEventListener('click', () => {
        els.pasteImportArea.style.display = 'none';
        els.pasteImportText.value = '';
    });
    els.pasteImportApply.addEventListener('click', () => {
        const text = els.pasteImportText.value.trim();
        if (!text) { showToast('Nothing to import', 'warning'); return; }
        try {
            const { columns, rows } = parseTabular(text);
            if (!rows.length) { showToast('No rows detected', 'warning'); return; }
            const append = els.pasteImportAppend.checked;
            if (append) {
                // Merge: union columns, preserve existing rows
                columns.forEach(c => { if (!state.manualColumns.includes(c)) state.manualColumns.push(c); });
                state.manualRows = state.manualRows.filter(r => Object.values(r).some(v => (v || '').toString().trim()));
                rows.forEach(r => {
                    const norm = {}; state.manualColumns.forEach(c => norm[c] = r[c] || '');
                    state.manualRows.push(norm);
                });
            } else {
                state.manualColumns = columns.slice();
                if (!state.manualColumns.includes('Email')) state.manualColumns.push('Email');
                if (!state.manualColumns.includes('Name')) state.manualColumns.unshift('Name');
                state.manualRows = rows.map(r => {
                    const norm = {}; state.manualColumns.forEach(c => norm[c] = r[c] || '');
                    return norm;
                });
            }
            state.selectedRows.clear();
            els.pasteImportArea.style.display = 'none';
            els.pasteImportText.value = '';
            renderManualGrid();
            refreshPlaceholders();
            showToast(`Imported ${rows.length} row(s)`, 'success');
        } catch (err) {
            showToast(`Import failed: ${err.message}`, 'error');
        }
    });

    // Export grid as CSV
    els.exportGridBtn.addEventListener('click', () => {
        const rows = state.manualRows.filter(r => Object.values(r).some(v => (v || '').toString().trim()));
        if (rows.length === 0) { showToast('Grid is empty', 'warning'); return; }
        const csv = toCSV(state.manualColumns, rows);
        downloadFile(csv, 'kiit_mailer_recipients.csv', 'text/csv');
        showToast(`Exported ${rows.length} row(s)`, 'success');
    });
}

// Tabular paste parser: detects tab-separated or comma-separated.
function parseTabular(text) {
    const lines = text.replace(/\r/g, '').split('\n').filter(l => l.length > 0);
    if (lines.length === 0) return { columns: [], rows: [] };
    const sep = lines[0].includes('\t') ? '\t' : ',';
    const splitLine = (line) => {
        if (sep === '\t') return line.split('\t').map(s => s.trim());
        // Lightweight CSV split (handles simple quoted fields)
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

function toCSV(columns, rows) {
    const esc = v => {
        const s = String(v == null ? '' : v);
        if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
        return s;
    };
    const head = columns.map(esc).join(',');
    const body = rows.map(r => columns.map(c => esc(r[c])).join(',')).join('\n');
    return head + '\n' + body + '\n';
}

function downloadFile(content, filename, mime) {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
}

function renderManualGrid() {
    const t = els.manualGrid;
    t.innerHTML = '';

    // Prune selectedRows of out-of-range indices
    state.selectedRows = new Set([...state.selectedRows].filter(i => i < state.manualRows.length));

    // Header
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');

    // Master checkbox
    const thCheck = document.createElement('th');
    thCheck.className = 'head-check';
    const allSelected = state.manualRows.length > 0 && state.selectedRows.size === state.manualRows.length;
    thCheck.innerHTML = `<input type="checkbox" id="masterCheck" ${allSelected ? 'checked' : ''} title="Select all rows">`;
    trh.appendChild(thCheck);

    state.manualColumns.forEach((col, ci) => {
        const th = document.createElement('th');
        const isRequired = col === 'Email';
        th.innerHTML = `
            <div class="th-inner">
                <input type="text" class="col-name" data-ci="${ci}" value="${escapeAttr(col)}" ${isRequired ? 'readonly title="Required column"' : ''}>
                ${isRequired ? '' : `<button type="button" class="icon-btn" data-delcol="${ci}" title="Delete column">✕</button>`}
            </div>
        `;
        trh.appendChild(th);
    });
    const thAct = document.createElement('th');
    thAct.style.width = '40px';
    trh.appendChild(thAct);
    thead.appendChild(trh);
    t.appendChild(thead);

    // Body
    const tbody = document.createElement('tbody');
    state.manualRows.forEach((row, ri) => {
        const tr = document.createElement('tr');
        if (state.selectedRows.has(ri)) tr.classList.add('selected');

        const checkTd = document.createElement('td');
        checkTd.className = 'row-check';
        checkTd.innerHTML = `<input type="checkbox" class="row-select" data-ri="${ri}" ${state.selectedRows.has(ri) ? 'checked' : ''}>`;
        tr.appendChild(checkTd);

        state.manualColumns.forEach(col => {
            const td = document.createElement('td');
            const inputType = col === 'Email' ? 'email' : 'text';
            td.innerHTML = `<input type="${inputType}" data-ri="${ri}" data-col="${escapeAttr(col)}" value="${escapeAttr(row[col] || '')}" placeholder="${escapeAttr(col)}">`;
            tr.appendChild(td);
        });
        const actTd = document.createElement('td');
        actTd.innerHTML = `<button type="button" class="icon-btn" data-delrow="${ri}" title="Delete row">✕</button>`;
        tr.appendChild(actTd);
        tbody.appendChild(tr);
    });
    t.appendChild(tbody);

    // Listeners
    const masterCheck = t.querySelector('#masterCheck');
    if (masterCheck) {
        masterCheck.addEventListener('change', e => {
            if (e.target.checked) {
                state.selectedRows = new Set(state.manualRows.map((_, i) => i));
            } else {
                state.selectedRows.clear();
            }
            renderManualGrid();
        });
    }
    t.querySelectorAll('.row-select').forEach(cb => {
        cb.addEventListener('change', e => {
            const ri = +e.target.dataset.ri;
            if (e.target.checked) state.selectedRows.add(ri);
            else state.selectedRows.delete(ri);
            renderManualGrid();
        });
    });
    t.querySelectorAll('input.col-name').forEach(inp => {
        inp.addEventListener('change', e => {
            const ci = +e.target.dataset.ci;
            const newName = e.target.value.trim();
            const oldName = state.manualColumns[ci];
            if (!newName) { e.target.value = oldName; return; }
            if (state.manualColumns.includes(newName) && newName !== oldName) {
                showToast('Column name must be unique', 'error');
                e.target.value = oldName; return;
            }
            state.manualColumns[ci] = newName;
            state.manualRows.forEach(r => { r[newName] = r[oldName] || ''; if (newName !== oldName) delete r[oldName]; });
            renderManualGrid();
            refreshPlaceholders();
        });
    });
    t.querySelectorAll('[data-delcol]').forEach(btn => {
        btn.addEventListener('click', e => {
            const ci = +e.currentTarget.dataset.delcol;
            const col = state.manualColumns[ci];
            state.manualColumns.splice(ci, 1);
            state.manualRows.forEach(r => delete r[col]);
            renderManualGrid();
            refreshPlaceholders();
        });
    });
    t.querySelectorAll('tbody input[type="text"], tbody input[type="email"]').forEach(inp => {
        inp.addEventListener('input', e => {
            const ri = +e.target.dataset.ri;
            const col = e.target.dataset.col;
            if (Number.isNaN(ri) || !col) return;
            state.manualRows[ri][col] = e.target.value;
            updateGridStatus();
        });
        // Paste-from-Excel handling: tabs/newlines distribute across grid
        inp.addEventListener('paste', e => {
            const text = (e.clipboardData || window.clipboardData).getData('text');
            if (!text) return;
            if (!/[\t\n]/.test(text)) return; // single-cell paste, leave default
            e.preventDefault();
            const startRi = +e.target.dataset.ri;
            const startCol = e.target.dataset.col;
            const startCi = state.manualColumns.indexOf(startCol);
            if (startCi < 0) return;
            const lines = text.replace(/\r/g, '').replace(/\n+$/, '').split('\n');
            lines.forEach((line, lineIdx) => {
                const cells = line.split('\t');
                const targetRi = startRi + lineIdx;
                while (state.manualRows.length <= targetRi) {
                    const empty = {}; state.manualColumns.forEach(c => empty[c] = '');
                    state.manualRows.push(empty);
                }
                cells.forEach((val, cellIdx) => {
                    const ci = startCi + cellIdx;
                    if (ci >= state.manualColumns.length) return;
                    const col = state.manualColumns[ci];
                    state.manualRows[targetRi][col] = val;
                });
            });
            renderManualGrid();
            refreshPlaceholders();
            showToast(`Pasted ${lines.length} row(s)`, 'success');
        });
    });
    t.querySelectorAll('[data-delrow]').forEach(btn => {
        btn.addEventListener('click', e => {
            const ri = +e.currentTarget.dataset.delrow;
            state.manualRows.splice(ri, 1);
            state.selectedRows.delete(ri);
            // shift indices in selectedRows
            state.selectedRows = new Set([...state.selectedRows].map(i => i > ri ? i - 1 : i));
            if (state.manualRows.length === 0) {
                const empty = {}; state.manualColumns.forEach(c => empty[c] = '');
                state.manualRows.push(empty);
            }
            renderManualGrid();
        });
    });

    updateGridStatus();
}

function updateGridStatus() {
    const valid = state.manualRows.filter(r => (r.Email || '').trim() !== '').length;
    els.gridStatus.textContent = `${state.manualRows.length} row(s) · ${valid} with email`;
    const sel = state.selectedRows.size;
    if (els.deleteSelectedBtn) {
        els.deleteSelectedBtn.style.display = sel > 0 ? 'inline-block' : 'none';
        if (els.selectedCount) els.selectedCount.textContent = sel;
    }
}

// ---------- Quill / Editors ----------
function initQuill() {
    state.quill = new Quill('#editor-container', {
        theme: 'snow',
        placeholder: 'Compose your email...',
        modules: {
            toolbar: [
                [{ header: [1, 2, 3, false] }],
                ['bold', 'italic', 'underline', 'strike'],
                [{ color: [] }, { background: [] }],
                [{ list: 'ordered' }, { list: 'bullet' }],
                [{ align: [] }],
                ['link', 'image', 'blockquote'],
                ['clean']
            ]
        }
    });

    els.htmlEditor.value = DEFAULT_HTML;
    els.textEditor.value = DEFAULT_TEXT;
}

function setupEditorMode() {
    els.editorSeg.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const newMode = btn.dataset.editor;
            if (newMode === state.editorMode) return;
            // Plain text vs HTML/Rich are independent; only warn between html<->rich
            const needsWarn = (state.editorMode === 'html' && newMode === 'rich') ||
                              (state.editorMode === 'rich' && newMode === 'html');
            if (needsWarn) {
                state.pendingEditorMode = newMode;
                els.modeWarningModal.classList.remove('hidden');
            } else {
                applyEditorMode(newMode, false);
            }
        });
    });

    els.confirmModeSwitch.addEventListener('click', () => {
        applyEditorMode(state.pendingEditorMode, true);
        els.modeWarningModal.classList.add('hidden');
    });
    els.cancelModeSwitch.addEventListener('click', () => {
        state.pendingEditorMode = null;
        els.modeWarningModal.classList.add('hidden');
    });
}

function applyEditorMode(mode, clearOther) {
    state.editorMode = mode;
    els.editorSeg.querySelectorAll('.seg-btn').forEach(b => b.classList.toggle('active', b.dataset.editor === mode));
    els.htmlEditorContainer.style.display = mode === 'html' ? 'block' : 'none';
    els.richEditorContainer.style.display = mode === 'rich' ? 'block' : 'none';
    els.textEditorContainer.style.display = mode === 'text' ? 'block' : 'none';

    if (clearOther) {
        if (mode === 'html') state.quill.setText('');
        else if (mode === 'rich') els.htmlEditor.value = '';
    }
}

// ---------- Placeholder Toolbar ----------
function refreshPlaceholders() {
    const headers = currentHeaders();
    els.placeholderToolbar.innerHTML = '';
    if (headers.length === 0) return;
    headers.forEach(h => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = `{${h}}`;
        btn.className = 'placeholder-btn';
        btn.title = `Insert {${h}} placeholder`;
        btn.onclick = () => insertPlaceholder(h);
        els.placeholderToolbar.appendChild(btn);
    });
}

function insertPlaceholder(h) {
    const token = `{${h}}`;
    if (state.editorMode === 'html') {
        insertAtCursor(els.htmlEditor, token);
    } else if (state.editorMode === 'text') {
        insertAtCursor(els.textEditor, token);
    } else {
        const range = state.quill.getSelection(true);
        state.quill.insertText(range ? range.index : 0, token);
    }
}

function insertAtCursor(textarea, text) {
    const start = textarea.selectionStart || 0;
    const end = textarea.selectionEnd || 0;
    textarea.value = textarea.value.slice(0, start) + text + textarea.value.slice(end);
    textarea.selectionStart = textarea.selectionEnd = start + text.length;
    textarea.focus();
}

function currentHeaders() {
    if (state.sendMode === 'single') {
        const set = new Set(['Name', 'Email']);
        state.singleExtras.forEach(f => { if (f.key && f.key.trim()) set.add(f.key.trim()); });
        return [...set];
    }
    if (state.bulkSource === 'csv') {
        return state.csvHeaders && state.csvHeaders.length ? state.csvHeaders : [];
    }
    return state.manualColumns.slice();
}

// ---------- Attachments ----------
function setupAttachments() {
    els.attachInput.addEventListener('change', e => {
        const newFiles = Array.from(e.target.files);
        let currentSize = state.attachments.reduce((acc, f) => acc + f.size, 0);
        const valid = [];
        let skipped = false;
        newFiles.forEach(f => {
            if (currentSize + f.size > MAX_SIZE_MB * 1024 * 1024) skipped = true;
            else { currentSize += f.size; valid.push(f); }
        });
        if (skipped) showToast(`Some files skipped — total exceeds ${MAX_SIZE_MB} MB`, 'warning');
        state.attachments = [...state.attachments, ...valid];
        renderAttachments();
        els.attachInput.value = '';
    });

    els.clearAttachmentsBtn.addEventListener('click', () => {
        state.attachments = [];
        renderAttachments();
    });
}

function renderAttachments() {
    els.attachList.innerHTML = '';
    els.clearAttachmentsBtn.style.display = state.attachments.length > 0 ? 'inline-block' : 'none';

    let total = 0;
    state.attachments.forEach((file, idx) => {
        total += file.size;
        const item = document.createElement('div');
        item.className = 'attach-chip';
        item.innerHTML = `
            <span>📎 ${file.name} <span class="muted-small">(${(file.size / 1024 / 1024).toFixed(2)} MB)</span></span>
            <button type="button" class="x" data-i="${idx}">×</button>
        `;
        item.querySelector('.x').addEventListener('click', () => {
            state.attachments.splice(idx, 1);
            renderAttachments();
        });
        els.attachList.appendChild(item);
    });

    if (total > 0) {
        const info = document.createElement('div');
        info.className = 'attach-total';
        info.style.color = total > MAX_SIZE_MB * 1024 * 1024 ? 'var(--danger)' : 'var(--text-muted)';
        info.textContent = `Total: ${(total / 1024 / 1024).toFixed(2)} MB / ${MAX_SIZE_MB} MB`;
        els.attachList.appendChild(info);
    }
}

// ---------- Preview ----------
function setupPreview() {
    els.previewBtn.addEventListener('click', openPreview);
    els.closePreviewBtn.addEventListener('click', () => els.previewModal.classList.add('hidden'));
    els.previewModal.addEventListener('click', e => {
        if (e.target === els.previewModal) els.previewModal.classList.add('hidden');
    });

    els.previewThemeSeg.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            els.previewThemeSeg.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.previewTheme = btn.dataset.theme;
            applyPreviewChrome();
            renderPreviewBody();
        });
    });
    els.previewViewportSeg.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            els.previewViewportSeg.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.previewViewport = btn.dataset.viewport;
            applyPreviewChrome();
        });
    });

    els.previewRecipient.addEventListener('change', e => {
        state.previewRowIndex = +e.target.value;
        renderPreviewBody();
    });
}

function getRecipients() {
    if (state.sendMode === 'single') {
        const row = { Name: els.singleName.value || '', Email: els.singleEmail.value || '' };
        state.singleExtras.forEach(f => { if (f.key) row[f.key] = f.value; });
        return [row];
    }
    if (state.bulkSource === 'csv') return state.csvData || [];
    return state.manualRows.filter(r => (r.Email || '').trim() !== '');
}

function openPreview() {
    const recipients = getRecipients();
    // Populate recipient dropdown
    els.previewRecipient.innerHTML = '';
    if (recipients.length > 1) {
        els.previewRecipient.style.display = 'inline-block';
        recipients.forEach((r, i) => {
            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = `${r.Name || '(no name)'} <${r.Email || 'no-email'}>`;
            els.previewRecipient.appendChild(opt);
        });
        state.previewRowIndex = 0;
    } else {
        els.previewRecipient.style.display = 'none';
        state.previewRowIndex = 0;
    }

    applyPreviewChrome();
    renderPreviewBody();
    els.previewModal.classList.remove('hidden');
}

function applyPreviewChrome() {
    els.gmailFrame.classList.toggle('dark', state.previewTheme === 'dark');
    els.gmailFrame.classList.toggle('mobile', state.previewViewport === 'mobile');
}

function renderPreviewBody() {
    const recipients = getRecipients();
    const row = recipients[state.previewRowIndex] || { Name: 'John Doe', Email: 'john@example.com' };

    // Subject + meta
    const creds = state.credentials || {};
    const fromName = creds.displayName || (creds.email ? creds.email.split('@')[0] : 'KIIT Mailer');
    els.gmSubject.textContent = substitute(els.subject.value || '(no subject)', row);
    els.gmFromName.textContent = fromName;
    els.gmFromEmail.textContent = creds.email ? `<${creds.email}>` : '<sender@example.com>';
    els.gmTo.textContent = row.Email || 'recipient@example.com';
    els.gmAvatar.textContent = (fromName || 'K').charAt(0).toUpperCase();
    els.gmDate.textContent = new Date().toLocaleString(undefined, { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });

    // HTML body
    let html = '';
    if (state.editorMode === 'html') html = els.htmlEditor.value;
    else if (state.editorMode === 'rich') html = wrapRichContent(state.quill.root.innerHTML);
    else html = textToHtml(els.textEditor.value);

    html = substitute(html, row);

    // Inject color-scheme to simulate theme inside iframe
    const themedHtml = injectColorScheme(html, state.previewTheme);
    els.previewFrame.onload = () => {
        try {
            const doc = els.previewFrame.contentDocument;
            if (!doc || !doc.body) return;
            // Measure with a brief delay so fonts/images settle
            const measure = () => {
                const h = Math.max(doc.body.scrollHeight, doc.documentElement.scrollHeight);
                els.previewFrame.style.height = Math.min(Math.max(h, 320), 2000) + 'px';
            };
            measure();
            setTimeout(measure, 120);
        } catch (e) { /* cross-origin: keep CSS height */ }
    };
    els.previewFrame.srcdoc = themedHtml;
}

function injectColorScheme(html, theme) {
    // Light mode: render as-is (clients honor light backgrounds).
    // Dark mode: simulate the *forced dark* algorithm that Gmail Android applies to
    // emails not explicitly designed for dark mode — invert the document, then
    // re-invert media so images/photos stay correct. Hue-rotate keeps colors.
    let injected;
    if (theme === 'dark') {
        injected = `<style id="__kiit_theme">
          :root { color-scheme: dark; }
          html { background: #1f1f1f !important; }
          html { filter: invert(1) hue-rotate(180deg); }
          img, picture, video, svg, iframe, canvas,
          [style*="background-image"], [data-skip-darken] {
            filter: invert(1) hue-rotate(180deg);
          }
        </style>`;
    } else {
        injected = `<style id="__kiit_theme">:root { color-scheme: light; }</style>`;
    }

    let out = html;
    if (/<\/head>/i.test(out)) {
        out = out.replace(/<\/head>/i, injected + '</head>');
    } else if (/<head[^>]*>/i.test(out)) {
        out = out.replace(/<head[^>]*>/i, m => m + injected);
    } else {
        out = injected + out;
    }
    return out;
}

function substitute(str, row) {
    if (!str) return str;
    return String(str).replace(/\{([^{}]+)\}/g, (_m, key) => {
        const v = row[key.trim()];
        return v != null ? v : `{${key}}`;
    });
}

function textToHtml(text) {
    const escaped = String(text || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
    const body = escaped.split(/\n/).map(l => l || '&nbsp;').join('<br>');
    return `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light"></head>
<body style="margin:0; padding:24px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif; font-size:14px; line-height:1.6; color:#1f2328; background:#ffffff;">
<div style="max-width:600px; margin:0 auto;">${body}</div>
</body></html>`;
}

function wrapRichContent(html) {
    if (/<!DOCTYPE/i.test(html)) return html;
    // Light-locked wrapper — opts out of mail-client dark-mode transformations
    // so the email appears exactly as the user composed it.
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
<div style="max-width:600px; margin:0 auto;">${html}</div>
</body>
</html>`;
}

// ---------- Sending ----------
function setupSending() {
    els.sendBtn.addEventListener('click', () => sendAll(false));
}

// Build current message payload (subject/html/text) from editor state.
function buildMessagePayload() {
    const subject = els.subject.value.trim();
    let htmlPayload = '';
    let textPayload = '';
    if (state.editorMode === 'html') {
        htmlPayload = els.htmlEditor.value;
    } else if (state.editorMode === 'rich') {
        htmlPayload = wrapRichContent(state.quill.root.innerHTML);
    } else {
        textPayload = els.textEditor.value || '';
    }
    // Preheader injection (HTML only)
    const preheader = (els.preheader.value || '').trim();
    if (preheader && htmlPayload) {
        htmlPayload = injectPreheader(htmlPayload, preheader);
    }
    return { subject, htmlPayload, textPayload, preheader };
}

function injectPreheader(html, preheader) {
    const safe = preheader
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const block = `<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;font-size:1px;line-height:1px;color:transparent;opacity:0;">${safe}</div>`;
    if (/<body[^>]*>/i.test(html)) {
        return html.replace(/<body[^>]*>/i, m => m + block);
    }
    return block + html;
}

// Build the resolved queue (per-recipient) given recipients and payload.
function buildQueue(recipients, payload) {
    return recipients.map((row, idx) => ({
        index: idx + 1,
        email: (row.Email || '').trim(),
        subject: substitute(payload.subject, row),
        html: payload.htmlPayload ? substitute(payload.htmlPayload, row) : '',
        text: payload.textPayload ? substitute(payload.textPayload, row) : '',
        row,
    }));
}

function dedupeRecipients(list) {
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

async function sendAll(isResume = false) {
    if (state.isSending) return;
    if (!state.credentials || !state.credentials.email || !state.credentials.pass) {
        showToast('Please configure settings first', 'error');
        els.settingsTabBtn.click();
        return;
    }

    let queue;
    if (isResume) {
        const saved = JSON.parse(localStorage.getItem(STORAGE_QUEUE) || 'null');
        if (!saved || !saved.queue || !saved.queue.length) {
            showToast('No saved queue to resume', 'warning');
            return;
        }
        queue = saved.queue;
        // Reset results table for the resumed run (prior results stay in logs/CSV downloads if user kept them).
        state.sendResults = [];
    } else {
        const payload = buildMessagePayload();
        if (!payload.subject) { showToast('Please enter a subject', 'warning'); return; }
        if (!payload.htmlPayload.trim() && !payload.textPayload.trim()) {
            showToast('Email body is empty', 'warning'); return;
        }

        let recipients = getRecipients();
        if (recipients.length === 0) {
            showToast(state.sendMode === 'single' ? 'Enter recipient email' : 'No recipients to send to', 'warning');
            return;
        }
        // Validate emails
        recipients = recipients.filter(r => /\S+@\S+\.\S+/.test((r.Email || '').trim()));
        if (recipients.length === 0) {
            showToast('No valid email addresses found', 'error'); return;
        }
        // Dedup (bulk only)
        if (state.sendMode === 'bulk' && els.dedupToggle && els.dedupToggle.checked) {
            const before = recipients.length;
            recipients = dedupeRecipients(recipients);
            const dropped = before - recipients.length;
            if (dropped > 0) log('system', `Deduplication dropped ${dropped} duplicate email(s)`);
        }

        queue = buildQueue(recipients, payload);
        state.sendResults = [];
    }

    state.sendQueue = queue;
    state.sendCursor = 0;
    state.isSending = true;
    state.isPaused = false;
    state.isCancelled = false;

    openSendConsole(queue.length);

    const delay = state.sendMode === 'bulk' ? Math.max(0, parseInt(els.delayMs.value, 10) || 0) : 0;

    log('system', `Starting batch of ${queue.length} email(s)${isResume ? ' (resumed)' : ''}...`);

    // Persist queue snapshot — used for resume after refresh/crash.
    persistQueue();

    while (state.sendCursor < state.sendQueue.length) {
        if (state.isCancelled) break;
        if (state.isPaused) {
            await sleep(200);
            continue;
        }
        const item = state.sendQueue[state.sendCursor];
        appendPendingResult(item);
        const result = await sendOne(item);
        finalizeResult(item, result);
        state.sendCursor++;
        updateConsoleProgress();
        persistQueue();

        if (delay > 0 && state.sendCursor < state.sendQueue.length && !state.isCancelled) {
            // Sleep but allow pause/cancel to interrupt promptly
            await sleepInterruptible(delay);
        }
    }

    finishSendConsole();
    if (!state.isCancelled) localStorage.removeItem(STORAGE_QUEUE);
    state.isSending = false;
}

async function sendOne(item) {
    const fd = new FormData();
    fd.append('to', item.email);
    fd.append('subject', item.subject);
    if (item.html) fd.append('html', item.html);
    if (item.text) fd.append('text', item.text);

    fd.append('smtpUser', state.credentials.email);
    fd.append('smtpPass', state.credentials.pass);
    if (state.credentials.displayName) fd.append('displayName', state.credentials.displayName);
    if (state.credentials.replyTo) fd.append('replyTo', state.credentials.replyTo);

    state.attachments.forEach(f => fd.append('file', f));

    try {
        const res = await fetch('/api/send-mail', { method: 'POST', body: fd });
        const raw = await res.text();
        let data;
        try { data = JSON.parse(raw); }
        catch (e) { throw new Error(`Server: ${raw.substring(0, 120)}`); }
        if (res.ok && data.success) {
            log('success', `Sent to ${item.email}`);
            return { ok: true, detail: data.messageId || 'OK' };
        } else {
            const err = data.error || 'Unknown error';
            log('error', `Failed → ${item.email}: ${err}`);
            return { ok: false, detail: err };
        }
    } catch (err) {
        log('error', `Error → ${item.email}: ${err.message}`);
        return { ok: false, detail: err.message };
    }
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
async function sleepInterruptible(ms) {
    const step = 100;
    let elapsed = 0;
    while (elapsed < ms) {
        if (state.isCancelled) return;
        await sleep(Math.min(step, ms - elapsed));
        elapsed += step;
        if (state.isPaused) {
            // Keep waiting in pause — but don't accumulate the delay
            while (state.isPaused && !state.isCancelled) await sleep(200);
        }
    }
}

function persistQueue() {
    if (!state.isSending && state.sendQueue.length === 0) return;
    const remaining = state.sendQueue.slice(state.sendCursor);
    if (remaining.length === 0) return;
    const snapshot = {
        savedAt: new Date().toISOString(),
        queue: remaining.map(q => ({
            index: q.index, email: q.email, subject: q.subject, html: q.html, text: q.text,
        })),
        results: state.sendResults,
    };
    localStorage.setItem(STORAGE_QUEUE, JSON.stringify(snapshot));
}

// ---------- Send Console UI ----------
function setupSendConsole() {
    els.pauseBtn.addEventListener('click', () => {
        state.isPaused = true;
        els.pauseBtn.style.display = 'none';
        els.resumeBtnConsole.style.display = 'inline-block';
        els.consoleTitle.textContent = 'Paused';
    });
    els.resumeBtnConsole.addEventListener('click', () => {
        state.isPaused = false;
        els.pauseBtn.style.display = 'inline-block';
        els.resumeBtnConsole.style.display = 'none';
        els.consoleTitle.textContent = 'Sending Emails…';
    });
    els.cancelBtn.addEventListener('click', () => {
        if (!confirm('Cancel sending? Remaining recipients will be skipped (queue saved for resume).')) return;
        state.isCancelled = true;
        state.isPaused = false;
    });
    els.exportResultsBtn.addEventListener('click', exportResultsCSV);
    els.closeConsoleBtn.addEventListener('click', () => {
        els.overlay.classList.add('hidden');
    });
}

function openSendConsole(total) {
    els.resultsTable.querySelector('tbody').innerHTML = '';
    els.successCount.textContent = '0';
    els.failureCount.textContent = '0';
    els.pendingCount.textContent = total;
    els.progressBar.style.width = '0%';
    els.progressText.textContent = `0 / ${total}`;
    els.progressPercent.textContent = '0%';
    els.consoleTitle.textContent = 'Sending Emails…';
    els.pauseBtn.style.display = 'inline-block';
    els.pauseBtn.disabled = false;
    els.resumeBtnConsole.style.display = 'none';
    els.cancelBtn.disabled = false;
    els.exportResultsBtn.disabled = true;
    els.closeConsoleBtn.style.display = 'none';
    els.overlay.classList.remove('hidden');
    els.sendBtn.disabled = true;
}

function appendPendingResult(item) {
    const tbody = els.resultsTable.querySelector('tbody');
    const tr = document.createElement('tr');
    tr.className = 'pending';
    tr.dataset.idx = item.index;
    tr.innerHTML = `<td>${item.index}</td><td>${escapeAttr(item.email)}</td><td class="status-cell">⏳ Sending…</td><td>—</td>`;
    tbody.appendChild(tr);
    tr.scrollIntoView({ block: 'nearest' });
}

function finalizeResult(item, result) {
    state.sendResults.push({ index: item.index, email: item.email, ok: result.ok, detail: result.detail });
    const tbody = els.resultsTable.querySelector('tbody');
    const tr = tbody.querySelector(`tr[data-idx="${item.index}"]`);
    if (tr) {
        tr.classList.remove('pending');
        tr.classList.add(result.ok ? 'success' : 'failure');
        tr.children[2].textContent = result.ok ? '✓ Sent' : '✕ Failed';
        tr.children[3].textContent = result.detail || '';
    }
}

function updateConsoleProgress() {
    const total = state.sendQueue.length;
    const done = state.sendCursor;
    const pct = total ? Math.round((done / total) * 100) : 0;
    els.progressBar.style.width = `${pct}%`;
    els.progressText.textContent = `${done} / ${total}`;
    els.progressPercent.textContent = `${pct}%`;
    const ok = state.sendResults.filter(r => r.ok).length;
    const fail = state.sendResults.filter(r => !r.ok).length;
    els.successCount.textContent = ok;
    els.failureCount.textContent = fail;
    els.pendingCount.textContent = Math.max(0, total - done);
}

function finishSendConsole() {
    const total = state.sendQueue.length;
    const ok = state.sendResults.filter(r => r.ok).length;
    const fail = state.sendResults.filter(r => !r.ok).length;
    els.consoleTitle.textContent = state.isCancelled
        ? `Cancelled — ${ok} sent, ${fail} failed, ${total - state.sendCursor} skipped`
        : `Done — ${ok} sent, ${fail} failed`;
    els.pauseBtn.style.display = 'none';
    els.resumeBtnConsole.style.display = 'none';
    els.cancelBtn.disabled = true;
    els.exportResultsBtn.disabled = state.sendResults.length === 0;
    els.closeConsoleBtn.style.display = 'inline-block';
    els.sendBtn.disabled = false;
    if (!state.isCancelled) {
        showToast(`Finished. Sent ${ok}/${total}`, ok === total ? 'success' : 'warning');
    } else {
        showToast(`Cancelled. ${ok} sent, ${total - state.sendCursor} skipped (queue saved)`, 'warning');
        checkPendingQueue();
    }
}

function exportResultsCSV() {
    if (state.sendResults.length === 0) return;
    const cols = ['Index', 'Email', 'Status', 'Detail'];
    const rows = state.sendResults.map(r => ({
        Index: r.index, Email: r.email, Status: r.ok ? 'Sent' : 'Failed', Detail: r.detail || ''
    }));
    downloadFile(toCSV(cols, rows), `kiit_mailer_results_${Date.now()}.csv`, 'text/csv');
}

// ---------- Test Send (to self) ----------
function setupTestSend() {
    els.sendTestBtn.addEventListener('click', sendTestToSelf);
}

async function sendTestToSelf() {
    if (state.isSending) return;
    if (!state.credentials || !state.credentials.email || !state.credentials.pass) {
        showToast('Please configure settings first', 'error');
        els.settingsTabBtn.click();
        return;
    }
    const payload = buildMessagePayload();
    if (!payload.subject) { showToast('Please enter a subject', 'warning'); return; }
    if (!payload.htmlPayload.trim() && !payload.textPayload.trim()) {
        showToast('Email body is empty', 'warning'); return;
    }

    // Use first row's data (or fake data) so placeholders resolve nicely.
    const recips = getRecipients().filter(r => Object.values(r).some(v => (v || '').toString().trim()));
    let row = recips[0];
    if (!row || !row.Email) {
        row = { Name: 'Test User', Email: state.credentials.email };
        currentHeaders().forEach(h => { if (!(h in row)) row[h] = `[${h}]`; });
    }
    // Force the To address to the user's own KIIT address
    row = { ...row, Email: state.credentials.email };

    const item = {
        index: 0,
        email: state.credentials.email,
        subject: '[TEST] ' + substitute(payload.subject, row),
        html: payload.htmlPayload ? substitute(payload.htmlPayload, row) : '',
        text: payload.textPayload ? substitute(payload.textPayload, row) : '',
    };

    els.sendTestBtn.disabled = true;
    showToast('Sending test to your inbox…', 'info');
    log('system', `Sending TEST to ${item.email}`);
    const result = await sendOne(item);
    els.sendTestBtn.disabled = false;
    if (result.ok) showToast('Test sent — check your inbox', 'success');
    else showToast(`Test failed: ${result.detail}`, 'error');
}

// ---------- Subject helpers ----------
function setupSubjectHelpers() {
    els.subject.addEventListener('input', updateSubjectHelpers);
    els.preheader.addEventListener('input', updatePreheaderCounter);
}

function updateSubjectHelpers() {
    const v = els.subject.value || '';
    const len = v.length;
    els.subjectCounter.textContent = `${len} chars`;
    els.subjectCounter.classList.toggle('warn', len > 60 && len <= 78);
    els.subjectCounter.classList.toggle('danger', len > 78);

    // Spam keyword detection (case-insensitive, whole-word-ish)
    const lower = v.toLowerCase();
    const hits = SPAM_WORDS.filter(w => lower.includes(w.toLowerCase()));
    if (hits.length > 0) {
        els.subjectWarn.style.display = 'block';
        els.subjectWarn.innerHTML = `⚠️ Possible spam-trigger word(s): ${hits.map(h => `<code>${escapeAttr(h)}</code>`).join(', ')}`;
    } else {
        els.subjectWarn.style.display = 'none';
    }
}

function updatePreheaderCounter() {
    const v = els.preheader.value || '';
    els.preheaderCounter.textContent = `${v.length} chars`;
    els.preheaderCounter.classList.toggle('warn', v.length > 0 && (v.length < 30 || v.length > 110));
}

// ---------- Bulk Options & Pre-flight ----------
function setupBulkOptions() {
    els.runPreflightBtn.addEventListener('click', runPreflight);
    els.refreshPreflightBtn.addEventListener('click', runPreflight);
}

function runPreflight() {
    const recipients = getRecipients();
    const total = recipients.length;
    const emailsRaw = recipients.map(r => (r.Email || '').trim()).filter(Boolean);
    const emailsLower = emailsRaw.map(e => e.toLowerCase());
    const unique = new Set(emailsLower);
    const invalid = emailsRaw.filter(e => !/^\S+@\S+\.\S+$/.test(e));
    const dupCount = emailsLower.length - unique.size;

    // Duplicate addresses (preserve first occurrence)
    const seen = new Set();
    const dupes = [];
    emailsLower.forEach(e => { if (seen.has(e)) dupes.push(e); else seen.add(e); });

    // Detect missing placeholders in body
    const payload = buildMessagePayload();
    const bodyText = (payload.subject || '') + '\n' + (payload.htmlPayload || '') + '\n' + (payload.textPayload || '');
    const tokensInBody = new Set();
    String(bodyText).replace(/\{([^{}]+)\}/g, (_m, k) => { tokensInBody.add(k.trim()); return _m; });
    const headers = currentHeaders();
    const missing = [...tokensInBody].filter(t => !headers.includes(t));

    const stats = [
        { label: 'Total recipients', value: total, kind: total > 0 ? 'ok' : 'warn' },
        { label: 'Unique emails', value: unique.size, kind: 'ok' },
        { label: 'Invalid emails', value: invalid.length, kind: invalid.length ? 'danger' : 'ok' },
        { label: 'Duplicate emails', value: dupCount, kind: dupCount ? 'warn' : 'ok' },
        { label: 'Missing placeholders', value: missing.length, kind: missing.length ? 'danger' : 'ok' },
    ];

    let html = stats.map(s => `
        <div class="preflight-stat ${s.kind}">
            <span class="label">${s.label}</span>
            <span class="value">${s.value}</span>
        </div>
    `).join('');

    if (invalid.length > 0) {
        html += `<div class="preflight-list"><strong>Invalid:</strong> ${invalid.slice(0, 10).map(e => `<span class="tag">${escapeAttr(e)}</span>`).join('')}${invalid.length > 10 ? ` …+${invalid.length - 10} more` : ''}</div>`;
    }
    if (dupes.length > 0) {
        const uniqueDupes = [...new Set(dupes)];
        html += `<div class="preflight-list"><strong>Duplicates:</strong> ${uniqueDupes.slice(0, 10).map(e => `<span class="tag warn">${escapeAttr(e)}</span>`).join('')}${uniqueDupes.length > 10 ? ` …+${uniqueDupes.length - 10} more` : ''}</div>`;
    }
    if (missing.length > 0) {
        html += `<div class="preflight-list"><strong>Body uses placeholders not in columns:</strong> ${missing.map(m => `<span class="tag">{${escapeAttr(m)}}</span>`).join('')}</div>`;
    }

    els.preflightBody.innerHTML = html;
    els.preflightPanel.style.display = 'block';
}

// ---------- Resume Queue ----------
function checkPendingQueue() {
    const saved = JSON.parse(localStorage.getItem(STORAGE_QUEUE) || 'null');
    if (!saved || !saved.queue || saved.queue.length === 0) {
        els.resumeBanner.style.display = 'none';
        return;
    }
    const when = saved.savedAt ? new Date(saved.savedAt).toLocaleString() : 'unknown';
    els.resumeMeta.textContent = `${saved.queue.length} unsent recipient(s) from ${when}`;
    els.resumeBanner.style.display = 'flex';
    els.resumeBtn.onclick = () => sendAll(true);
    els.discardResumeBtn.onclick = () => {
        if (!confirm('Discard the saved unfinished batch?')) return;
        localStorage.removeItem(STORAGE_QUEUE);
        els.resumeBanner.style.display = 'none';
        showToast('Saved queue discarded', 'success');
    };
}

// ---------- Toasts & Logs ----------
function showToast(msg, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    const icon = { info: 'ℹ️', success: '✅', error: '❌', warning: '⚠️' }[type] || 'ℹ️';
    toast.innerHTML = `<span style="font-size:1.2em">${icon}</span> <span>${msg}</span>`;
    els.toastContainer.appendChild(toast);
    setTimeout(() => {
        toast.style.animation = 'slideIn 0.3s reverse';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

function log(type, msg) {
    if (els.logTerminal) {
        const t = new Date().toLocaleTimeString();
        const div = document.createElement('div');
        div.className = `log-line ${type}`;
        div.textContent = `[${t}] ${msg}`;
        els.logTerminal.appendChild(div);
        els.logTerminal.scrollTop = els.logTerminal.scrollHeight;
    }
    const logs = JSON.parse(localStorage.getItem(STORAGE_LOGS) || '[]');
    logs.push({ type, msg, time: new Date().toLocaleTimeString() });
    if (logs.length > 200) logs.shift();
    localStorage.setItem(STORAGE_LOGS, JSON.stringify(logs));
}

function loadHistory() {
    const logs = JSON.parse(localStorage.getItem(STORAGE_LOGS) || '[]');
    if (!els.logTerminal) return;
    els.logTerminal.innerHTML = '<div class="log-line system">System ready.</div>';
    logs.forEach(l => {
        const div = document.createElement('div');
        div.className = `log-line ${l.type}`;
        div.textContent = `[${l.time}] ${l.msg}`;
        els.logTerminal.appendChild(div);
    });
    els.logTerminal.scrollTop = els.logTerminal.scrollHeight;
}

function setupLogs() {
    if (!els.clearLogsBtn) return;
    els.clearLogsBtn.addEventListener('click', () => {
        els.logTerminal.innerHTML = '<div class="log-line system">Local logs cleared.</div>';
        localStorage.removeItem(STORAGE_LOGS);
    });
}
