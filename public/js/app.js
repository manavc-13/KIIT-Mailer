// KIIT Mailer - Frontend Logic
// HTML-first email composer with Single & Bulk (CSV / Manual Grid) modes,
// Gmail-style preview with theme + viewport toggles.

const MAX_SIZE_MB = 25;
const STORAGE_SETTINGS = 'kiit_mailer_settings';
const STORAGE_LOGS = 'kiit_mailer_logs';

// ---------- Default Starter HTML (generic, KIIT-friendly, light+dark safe) ----------
const DEFAULT_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light dark">
<meta name="supported-color-schemes" content="light dark">
<title>KIIT Mailer</title>
<style>
  body { margin:0; padding:0; -webkit-text-size-adjust:100%; }
  @media (prefers-color-scheme: dark) {
    .body-bg { background-color:#0f1115 !important; }
    .card { background-color:#1a1d24 !important; color:#e6e8eb !important; }
    .muted { color:#a8adb6 !important; }
    .accent { color:#7dabf8 !important; }
    .divider { border-top-color:#2a2e36 !important; }
    .btn { background-color:#2563eb !important; color:#ffffff !important; }
  }
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

    // Common
    subject: document.getElementById('subject'),
    placeholderToolbar: document.getElementById('placeholderToolbar'),
    attachInput: document.getElementById('attachments'),
    attachList: document.getElementById('attachmentList'),
    clearAttachmentsBtn: document.getElementById('clearAttachmentsBtn'),
    sendBtn: document.getElementById('sendBtn'),

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
    quill: null,
    credentials: null,

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

    renderManualGrid();
    refreshPlaceholders();
    loadHistory();

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
        renderManualGrid();
        refreshPlaceholders();
    });
}

function renderManualGrid() {
    const t = els.manualGrid;
    t.innerHTML = '';

    // Header
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
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
    t.querySelectorAll('tbody input').forEach(inp => {
        inp.addEventListener('input', e => {
            const ri = +e.target.dataset.ri;
            const col = e.target.dataset.col;
            state.manualRows[ri][col] = e.target.value;
            updateGridStatus();
        });
    });
    t.querySelectorAll('[data-delrow]').forEach(btn => {
        btn.addEventListener('click', e => {
            const ri = +e.currentTarget.dataset.delrow;
            state.manualRows.splice(ri, 1);
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
    return `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="color-scheme" content="light dark"></head>
<body style="margin:0; padding:24px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif; font-size:14px; line-height:1.6; color:#1f2328; background:#ffffff;">
<div style="max-width:600px; margin:0 auto;">${body}</div>
</body></html>`;
}

function wrapRichContent(html) {
    if (/<!DOCTYPE/i.test(html)) return html;
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light dark">
<style>
  body { margin:0; padding:24px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif; color:#1f2328; background:#ffffff; }
  @media (prefers-color-scheme: dark) {
    body { background:#0f1115 !important; color:#e6e8eb !important; }
  }
</style>
</head>
<body>
<div style="max-width:600px; margin:0 auto;">${html}</div>
</body>
</html>`;
}

// ---------- Sending ----------
function setupSending() {
    els.sendBtn.addEventListener('click', sendAll);
}

async function sendAll() {
    if (state.isSending) return;
    if (!state.credentials || !state.credentials.email || !state.credentials.pass) {
        showToast('Please configure settings first', 'error');
        els.settingsTabBtn.click();
        return;
    }

    const subject = els.subject.value.trim();
    if (!subject) { showToast('Please enter a subject', 'warning'); return; }

    // Build payload — only the active editor's content is sent
    let htmlPayload = '';
    let textPayload = '';
    if (state.editorMode === 'html') {
        htmlPayload = els.htmlEditor.value;
    } else if (state.editorMode === 'rich') {
        htmlPayload = wrapRichContent(state.quill.root.innerHTML);
    } else {
        textPayload = els.textEditor.value || '';
    }

    if (!htmlPayload.trim() && !textPayload.trim()) {
        showToast('Email body is empty', 'warning'); return;
    }

    const recipients = getRecipients();
    if (recipients.length === 0) {
        showToast(state.sendMode === 'single' ? 'Enter recipient email' : 'No recipients to send to', 'warning');
        return;
    }
    // Validate emails
    const valid = recipients.filter(r => /\S+@\S+\.\S+/.test((r.Email || '').trim()));
    if (valid.length === 0) {
        showToast('No valid email addresses found', 'error'); return;
    }

    const total = valid.length;
    startProgress(total);
    let success = 0;

    for (let i = 0; i < total; i++) {
        const row = valid[i];
        const email = row.Email.trim();
        const html = substitute(htmlPayload, row);
        const text = textPayload ? substitute(textPayload, row) : '';
        const subj = substitute(subject, row);
        const ok = await sendOne(email, subj, html, text, i + 1, total);
        if (ok) success++;
    }

    endProgress();
    showToast(`Finished. Sent ${success}/${total}`, success === total ? 'success' : 'warning');
}

async function sendOne(to, subject, html, text, index, total) {
    updateProgress(index, total);

    const fd = new FormData();
    fd.append('to', to);
    fd.append('subject', subject);
    if (html) fd.append('html', html);
    if (text) fd.append('text', text);

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
            log('success', `Sent to ${to}`);
            return true;
        } else {
            log('error', `Failed → ${to}: ${data.error || 'Unknown error'}`);
            return false;
        }
    } catch (err) {
        log('error', `Error → ${to}: ${err.message}`);
        return false;
    }
}

function startProgress(total) {
    state.isSending = true;
    els.overlay.classList.remove('hidden');
    els.sendBtn.disabled = true;
    log('system', `Starting batch of ${total} email(s)...`);
}
function updateProgress(current, total) {
    const pct = Math.round((current / total) * 100);
    els.progressBar.style.width = `${pct}%`;
    els.progressText.textContent = `${current} / ${total}`;
    els.progressPercent.textContent = `${pct}%`;
}
function endProgress() {
    state.isSending = false;
    setTimeout(() => {
        els.overlay.classList.add('hidden');
        els.sendBtn.disabled = false;
    }, 400);
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
