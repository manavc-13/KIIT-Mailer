
// Global UI State
const MAX_SIZE_MB = 25;

// DOM Elements
const els = {
    tabs: document.querySelectorAll('.nav-btn'),
    contents: document.querySelectorAll('.tab-content'),
    csvGroup: document.getElementById('csv-group'),
    csvFile: document.getElementById('csvFile'),
    csvStatus: document.getElementById('csvStatus'),
    placeholderToolbar: document.getElementById('placeholderToolbar'),
    attachInput: document.getElementById('attachments'),
    attachList: document.getElementById('attachmentList'),
    downloadTmplBtn: document.getElementById('downloadTemplateBtn'),
    sendBtn: document.getElementById('sendBtn'),
    // Logs
    logTerminal: document.getElementById('logTerminal'),
    clearLogsBtn: document.getElementById('clearLogs'),
    // Progress
    overlay: document.getElementById('progressOverlay'),
    progressBar: document.getElementById('progressBar'),
    progressText: document.getElementById('progressText'),
    progressPercent: document.getElementById('progressPercent'),
    subject: document.getElementById('subject'),
    logoutBtn: document.querySelector('.logout'),

    userNameDisplay: document.getElementById('userNameDisplay'),
    userAvatar: document.getElementById('userAvatar'),
    toastContainer: document.getElementById('toastContainer'),
    // Settings Elements
    settingsEmail: document.getElementById('settingsEmail'),
    settingsPass: document.getElementById('settingsPass'),
    settingsDisplayName: document.getElementById('settingsDisplayName'),
    settingsReplyTo: document.getElementById('settingsReplyTo'),
    saveSettingsBtn: document.getElementById('saveSettingsBtn'),
    settingsTabBtn: document.getElementById('settingsTabBtn'),
    // Editor Elements
    htmlEditor: document.getElementById('htmlEditor'),
    richEditorContainer: document.getElementById('richEditorParams'),
    htmlEditorContainer: document.getElementById('htmlEditorParams'),
    editorToggles: document.getElementsByName('editorMode'),
    previewBtn: document.getElementById('previewBtn'),
    previewModal: document.getElementById('previewModal'),
    closePreviewBtn: document.getElementById('closePreviewBtn'),
    previewFrame: document.getElementById('previewFrame'),
    // Warnings
    modeWarningModal: document.getElementById('modeWarningModal'),
    confirmModeSwitch: document.getElementById('confirmModeSwitch'),
    cancelModeSwitch: document.getElementById('cancelModeSwitch'),
    // Attachments
    clearAttachmentsBtn: document.getElementById('clearAttachmentsBtn')
};

// State
let state = {
    mode: 'bulk', // Forced bulk
    editorMode: 'html', // 'html' or 'rich'
    pendingEditorMode: null, // For warning
    attachments: [],
    csvData: null,
    csvHeaders: [],
    isSending: false,
    quill: null,
    credentials: null,
    fontLink: '',
    styleBlock: ''
};

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    // 0. Set Copyright Year
    document.getElementById('year').textContent = new Date().getFullYear();

    // 1. Load Settings
    loadSettings();

    // 2. Setup Quill
    initQuill();

    // 3. Setup Listeners
    setupTabs();
    setupCSV();
    setupAttachments();
    setupSending();
    setupLogs();
    setupSettings();
    setupSettings(); // Duplicate? Keeping one
    setupEditorMode(); // New Listener setup

    loadDefaultAttachment(); // Load default attachment
    loadHistory(); // Restore logs

    // Check if settings are missing, if so, switch to settings tab
    if (!state.credentials || !state.credentials.email || !state.credentials.pass) {
        showToast('Please configure settings first', 'warning');
        setTimeout(() => els.settingsTabBtn.click(), 500);
    }
});

// --- Settings ---
function loadSettings() {
    const stored = localStorage.getItem('kiit_mailer_settings');
    if (stored) {
        try {
            state.credentials = JSON.parse(stored);
            // Fill UI
            els.settingsEmail.value = state.credentials.email || '';
            els.settingsPass.value = state.credentials.pass || '';
            els.settingsDisplayName.value = state.credentials.displayName || '';

            // Ensure Reply-To has a value, fallback to default if missing in saved data
            els.settingsReplyTo.value = state.credentials.replyTo || 'info@edgei.org';
            els.settingsDisplayName.value = state.credentials.displayName || 'EDGEI 2026';

            // Update Profile UI
            if (state.credentials.email) {
                els.userNameDisplay.textContent = state.credentials.displayName || state.credentials.email.split('@')[0];
                els.userAvatar.textContent = (state.credentials.displayName || state.credentials.email).charAt(0).toUpperCase();
            }
        } catch (e) {
            console.error("Failed to parse settings", e);
        }
    } else {
        // No stored settings, set default reply-to
        // Note: we don't save it yet until user clicks save, but we show it in UI
        els.settingsReplyTo.value = 'info@edgei.org';
        els.settingsDisplayName.value = 'EDGEI 2026';
    }
}

// --- Default Attachment ---
function loadDefaultAttachment() {
    fetch('images/Edgei.jpeg')
        .then(response => {
            if (!response.ok) throw new Error("Image not found");
            return response.blob();
        })
        .then(blob => {
            const file = new File([blob], 'Edgei.jpeg', { type: 'image/jpeg' });
            state.attachments.push(file);
            renderAttachments();
            showToast('Default attachment loaded', 'success');
        })
        .catch(err => console.log('Default attachment not found or skipped:', err));
}

function setupSettings() {
    els.saveSettingsBtn.addEventListener('click', () => {
        const email = els.settingsEmail.value.trim();
        const pass = els.settingsPass.value.trim();
        const displayName = els.settingsDisplayName.value.trim();
        const replyTo = els.settingsReplyTo.value.trim() || 'info@edgei.org';

        // Validation
        if (!email.endsWith('@kiit.ac.in')) {
            showToast('Email must be @kiit.ac.in', 'error');
            return;
        }

        if (!pass) {
            showToast('App Password is required', 'error');
            return;
        }

        const creds = { email, pass, displayName, replyTo };
        localStorage.setItem('kiit_mailer_settings', JSON.stringify(creds));
        state.credentials = creds;

        // Update Profile
        els.userNameDisplay.textContent = displayName || email.split('@')[0];
        els.userAvatar.textContent = (displayName || email).charAt(0).toUpperCase();

        showToast('Settings Saved', 'success');
    });
}


// --- Quill Editor ---
function initQuill() {
    const Font = Quill.import('formats/font');
    Font.whitelist = ['inter', 'roboto', 'lato', 'montserrat', 'oswald', 'merriweather'];
    Quill.register(Font, true);

    state.quill = new Quill('#editor-container', {
        theme: 'snow',
        placeholder: 'Compose your email...',
        modules: {
            toolbar: [
                [{ 'header': [1, 2, 3, false] }],
                [{ 'font': Font.whitelist }],
                ['bold', 'italic', 'underline', 'strike'],
                [{ 'color': [] }, { 'background': [] }],
                [{ 'list': 'ordered' }, { 'list': 'bullet' }],
                ['link', 'image', 'clean']
            ]
        }
    });

    // Default Content & Styles
    const fontLink = '<link href="https://fonts.googleapis.com/css2?family=EB+Garamond:wght@400;500;600;700;800&display=swap" rel="stylesheet">';
    const styleBlock = `<style>
    /* Mobile Optimization */
    @media only screen and (max-width: 600px) {
        .main-table { width: 100% !important; max-width: 100% !important; }
        .mobile-padding { padding-left: 20px !important; padding-right: 20px !important; }
        .stack-column { display: block !important; width: 100% !important; padding-right: 0 !important; padding-bottom: 20px !important; }
        .stack-column-last { display: block !important; width: 100% !important; border-left: none !important; border-top: 1px solid #eeeeee !important; padding-top: 20px !important; padding-left: 0 !important; }
        .mobile-button { width: 100% !important; display: block !important; box-sizing: border-box !important; }
    }
</style>`;

    const defaultSubject = 'Invited Call For Papers: EDGEi-2026, Malaysia || Springer Series (Scopus)';

    // We construct the full HTML for the clipboard (Quill will parse what it can)
    // We also store this to wrap the outgoing mail later
    const fullDefaultHTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light dark">
<meta name="supported-color-schemes" content="light dark">
<title>EDGEI-2026 Invitation</title>
<link href="https://fonts.googleapis.com/css2?family=EB+Garamond:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    /* Global Resets */
    body { margin: 0; padding: 0; width: 100% !important; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
    
    /* Dark Mode Overrides (Works in Apple Mail, Outlook, Gmail Web) */
    @media (prefers-color-scheme: dark) {
        .body-bg { background-color: #121212 !important; }
        .content-bg { background-color: #121212 !important; color: #e0e0e0 !important; }
        .text-primary { color: #e0e0e0 !important; }
        .text-secondary { color: #bbbbbb !important; }
        .header-blue { color: #6fa8dc !important; } /* Lighter blue for dark mode */
        .link-blue { color: #6fa8dc !important; }
        .box-light { background-color: #1e1e1e !important; border-color: #444444 !important; }
        .divider { border-top: 1px solid #444444 !important; }
        .border-color { border-color: #444444 !important; }
        .schedule-header { background-color: #2c2c2c !important; color: #ffffff !important; }
        .deadline-box { background-color: #2a1515 !important; border-color: #5c2b2b !important; }
        .deadline-text { color: #ff9999 !important; }
        .footer-bg { background-color: #121212 !important; border-top: 1px solid #333333 !important; }
    }

    /* Mobile Optimization */
    @media only screen and (max-width: 600px) {
        .main-table { width: 100% !important; max-width: 100% !important; }
        .mobile-padding { padding-left: 20px !important; padding-right: 20px !important; }
        .stack-column { display: block !important; width: 100% !important; padding-right: 0 !important; padding-bottom: 20px !important; border-right: none !important; }
        .stack-column-last { display: block !important; width: 100% !important; border-left: none !important; border-top: 1px solid #eeeeee !important; padding-top: 20px !important; padding-left: 0 !important; }
        /* Dark Mode Mobile Border Fix */
        @media (prefers-color-scheme: dark) {
           .stack-column-last { border-top: 1px solid #444444 !important; }
        }
        .mobile-button { width: 100% !important; display: block !important; box-sizing: border-box !important; }
    }
</style>
</head>

<body class="body-bg" style="margin:0; padding:0; background-color:#ffffff; color:#222222;">

<table class="body-bg" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff;">
<tr>
<td align="center" style="padding: 20px 0;">

<table class="main-table content-bg" width="600" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff; max-width:600px; width:100%; margin:0 auto;">

    <tr>
        <td style="padding:30px 44px 0 44px;" class="mobile-padding">
            <h1 class="header-blue" style="
                margin:0 0 10px 0;
                font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                font-size:26px;
                font-weight:800;
                color:#0b5394;
                line-height: 1.1;
            ">
                International Conference on<br>Edge Intelligence (EDGEI-2026)
            </h1>

            <p class="text-primary" style="
                margin:0 0 12px 0;
                font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                font-size:16px;
                font-weight:500;
                color:#333333;
                line-height: 1.4;
            ">
                <strong>16-17 May, 2026</strong><br>
                Proceedings by <strong>Springer</strong> | Indexed by <strong>SCOPUS</strong><br>
                <a href="https://edgei.org" class="link-blue" style="color:#0b5394; text-decoration:underline;">www.edgei.org</a>
            </p>

            <div class="box-light" style="
                background-color: #f7f9fc;
                border-left: 4px solid #0b5394;
                padding: 10px 15px;
                margin-bottom: 5px;
            ">
                <p class="text-secondary" style="
                    margin:0;
                    font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                    font-size:16px;
                    color:#555555;
                    line-height:1.4;
                ">
                    <strong>Organizer:</strong><br>
                    Spectrum International University College, Malaysia
                </p>
            </div>
        </td>
    </tr>

    <tr>
        <td style="padding:20px 44px 0 44px;" class="mobile-padding">
            <hr class="divider" style="border:none; border-top:1px solid #eeeeee; margin:0;">
        </td>
    </tr>

    <tr>
        <td style="padding:24px 44px;" class="mobile-padding">
            <p class="text-primary" style="
                margin:0 0 16px 0;
                font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                font-size:18px;
                font-weight:500;
                color:#222222;
                line-height:1.6;
            ">
                Dear {Name},
            </p>

            <p class="text-primary" style="
                margin:0 0 16px 0;
                font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                font-size:18px;
                font-weight:500;
                color:#222222;
                line-height:1.6;
            ">
                The <strong>EDGEI-2026</strong> conference will be conducted in <strong>Hybrid mode</strong>. We invite you to submit your research to this premier forum.
            </p>
        </td>
    </tr>

    <tr>
        <td style="padding:0 44px 24px 44px;" class="mobile-padding">
            <table class="border-color" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #e0e0e0; border-radius: 6px; overflow: hidden;">
                <tr>
                    <td colspan="2" class="schedule-header" style="background-color:#f0f4f8; padding:10px 20px; border-bottom:1px solid #e0e0e0;">
                        <p class="schedule-header" style="margin:0; font-family:'EB Garamond', serif; font-size:16px; font-weight:700; color:#0b5394;">
                            Hybrid Schedule
                        </p>
                    </td>
                </tr>
                <tr>
                    <td width="100%" style="padding:0;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0">
                            <tr>
                                <td class="stack-column border-color" width="50%" valign="top" style="padding:20px; border-right:1px solid #e0e0e0;">
                                    <p class="text-primary" style="margin:0 0 5px 0; font-family:'EB Garamond', serif; font-size:18px; font-weight:700; color:#222222;">
                                        Physical Mode
                                    </p>
                                    <p class="text-secondary" style="margin:0; font-family:'EB Garamond', serif; font-size:16px; color:#444444; line-height:1.4;">
                                        May 16, 2026 (Wednesday)<br>
                                        Spectrum International University College, Malaysia
                                    </p>
                                </td>
                                <td class="stack-column-last" width="50%" valign="top" style="padding:20px;">
                                    <p class="text-primary" style="margin:0 0 5px 0; font-family:'EB Garamond', serif; font-size:18px; font-weight:700; color:#222222;">
                                        Online Mode
                                    </p>
                                    <p class="text-secondary" style="margin:0; font-family:'EB Garamond', serif; font-size:16px; color:#444444; line-height:1.4;">
                                        May 17, 2026 (Thursday)<br>
                                        Google Meet Platform
                                    </p>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>

    <tr>
        <td style="padding:0 44px 10px 44px;" class="mobile-padding">
             <p class="text-primary" style="
                margin:0 0 10px 0;
                font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                font-size:18px;
                color:#222222;
                line-height:1.6;
            ">
                <strong>Publication & Indexing:</strong><br>
                All EDGEI-2026 registered and presented papers will be published in conference proceedings by <strong>Springer</strong>. Papers published in books under this series are indexed by <strong>SCOPUS</strong>, etc.
            </p>
            <p class="text-primary" style="
                margin:0;
                font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                font-size:18px;
                color:#222222;
                line-height:1.6;
            ">
                <strong>Topics of Interest:</strong><br>
                Submissions of high-quality papers are expected in all areas of research and application in <strong>intelligent manufacturing and energy</strong>.
            </p>
        </td>
    </tr>

    <tr>
        <td style="padding:20px 44px 28px 44px;" class="mobile-padding">
            <div class="deadline-box" style="background-color:#fff5f5; border:1px dashed #d1a3a3; padding:12px; border-radius:4px; text-align:center;">
                <p class="deadline-text" style="margin:0; font-family:Arial, Helvetica, sans-serif; font-size:15px; font-weight:700; color:#a61c00;">
                    Paper Submission Deadline: 10 March 2026
                </p>
            </div>
        </td>
    </tr>

    <tr>
        <td align="center" style="padding:0 44px 32px 44px;" class="mobile-padding">
            <table border="0" cellspacing="0" cellpadding="0" style="margin: 0 0 15px 0;">
                <tr>
                    <td align="center" bgcolor="#0b5394" style="border-radius: 4px;">
                        <a href="https://cmt3.research.microsoft.com/User/Login?ReturnUrl=%2Fedgei2026" target="_blank" class="mobile-button" style="
                            font-size: 16px;
                            font-family: Arial, Helvetica, sans-serif;
                            color: #ffffff;
                            text-decoration: none;
                            padding: 14px 30px;
                            border: 1px solid #0b5394;
                            display: inline-block;
                            font-weight: 700;
                            border-radius: 4px;
                        ">
                            Submit Paper via Microsoft CMT
                        </a>
                    </td>
                </tr>
            </table>

            <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td align="center" style="border-radius: 4px;">
                        <a href="https://edgei.org" target="_blank" class="mobile-button link-blue" style="
                            font-size: 15px;
                            font-family: Arial, Helvetica, sans-serif;
                            color: #0b5394;
                            text-decoration: none;
                            padding: 12px 24px;
                            border: 2px solid #0b5394;
                            display: inline-block;
                            font-weight: 700;
                            border-radius: 4px;
                        ">
                            Visit Conference Website
                        </a>
                    </td>
                </tr>
            </table>
        </td>
    </tr>

    <tr>
        <td class="footer-bg" style="padding:24px 44px 36px 44px; border-top:1px solid #eeeeee; background-color: #ffffff;" class="mobile-padding">
            <p class="text-primary" style="
                margin:0 0 5px 0;
                font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                font-size:18px;
                color:#222222;
                font-weight: 500;
            ">
                Sincerely Yours,
            </p>
            <p class="text-secondary" style="
                margin:0;
                font-family:'EB Garamond', Garamond, 'Times New Roman', serif;
                font-size:17px;
                color:#444444;
                line-height:1.5;
            ">
                On behalf of Organizing Chair<br>
                <strong>EDGEI-2026</strong><br>
                <span style="font-size:15px;">Email: <a href="mailto:info@edgei.org" class="link-blue" style="color:#0b5394; text-decoration:none;">info@edgei.org</a></span>
            </p>
        </td>
    </tr>

</table>
</td>
</tr>
</table>
</body>
</html>`;

    // Make state global accessible for wrapper
    state.fontLink = fontLink;
    state.styleBlock = styleBlock;

    // Inject styles into the current document head so the editor renders the font
    document.head.insertAdjacentHTML('beforeend', fontLink);

    // Initial Default Content Load
    els.htmlEditor.value = fullDefaultHTML;
    // We also load it into Quill just in case (optional, but good for sync)
    // Quill might strip some things, but that's expected for Rich Text mode
    state.quill.clipboard.dangerouslyPasteHTML(fullDefaultHTML);

    // Set Subject
    if (!els.subject.value) {
        els.subject.value = defaultSubject;
    }
}

// --- Editor Mode & Preview ---
function setupEditorMode() {
    // Toggle Logic with Warning
    els.editorToggles.forEach(radio => {
        // Find the label wrapper to add click event, or just handle change
        // We need to intercept the change if possible, or rollback.
        // Easier: uncheck the new one if cancelled? 
        // Better: Make custom buttons, but radio is standard.
        // Let's rely on 'click' to prevent default if needed, but 'change' is safer for state.

        radio.addEventListener('click', (e) => {
            const newMode = e.target.value;
            if (newMode === state.editorMode) return; // No change

            // Prevent immediate switch
            e.preventDefault();

            state.pendingEditorMode = newMode;
            els.modeWarningModal.classList.remove('hidden');
        });
    });

    // Warning Modal Actions
    els.confirmModeSwitch.addEventListener('click', () => {
        state.editorMode = state.pendingEditorMode;

        // Manually check the radio button (since we prevented default)
        els.editorToggles.forEach(r => r.checked = (r.value === state.editorMode));

        updateEditorVisibility(true); // true = clear other
        els.modeWarningModal.classList.add('hidden');
    });

    els.cancelModeSwitch.addEventListener('click', () => {
        state.pendingEditorMode = null;
        els.modeWarningModal.classList.add('hidden');
    });

    // Preview Logic
    els.previewBtn.addEventListener('click', showPreview);
    els.closePreviewBtn.addEventListener('click', () => {
        els.previewModal.classList.add('hidden');
    });

    // Close preview on clicking outside
    els.previewModal.addEventListener('click', (e) => {
        if (e.target === els.previewModal) {
            els.previewModal.classList.add('hidden');
        }
    });
}

function updateEditorVisibility(clearOther = false) {
    if (state.editorMode === 'html') {
        els.htmlEditorContainer.style.display = 'block';
        els.richEditorContainer.style.display = 'none';

        if (clearOther) {
            state.quill.setText(''); // Clear Rich Text
            showToast('Rich Text Editor cleared for safety.', 'info');
        }
    } else {
        els.htmlEditorContainer.style.display = 'none';
        els.richEditorContainer.style.display = 'block';

        if (clearOther) {
            els.htmlEditor.value = ''; // Clear HTML
            showToast('HTML Editor cleared for safety.', 'info');
        }
    }
}

function showPreview() {
    let content = '';
    if (state.editorMode === 'html') {
        content = els.htmlEditor.value;
    } else {
        // Use the same wrapper logic as sending
        content = wrapRichContent(state.quill.root.innerHTML);
    }

    // Mock variable replacement
    // If CSV data exists, use the first row. Else use placeholders.
    if (state.csvData && state.csvData.length > 0) {
        const row = state.csvData[0];
        state.csvHeaders.forEach(h => {
            const reg = new RegExp(`{${h}}`, 'g');
            content = content.replace(reg, row[h] || `[${h}]`);
        });
    } else {
        // Just generic replace for preview
        content = content.replace(/{Name}/g, "John Doe");
        content = content.replace(/{Email}/g, "john@example.com");
    }

    els.previewFrame.srcdoc = content;
    els.previewModal.classList.remove('hidden');
}

function wrapRichContent(html) {
    if (!html.includes('<!DOCTYPE html>')) {
        return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
${state.fontLink || ''}
${state.styleBlock || ''}
</head>
<body style="margin:0; padding:0; background-color:#f4f4f4;">
${html}
</body>
</html>`;
    }
    return html;
}



// --- Toasts & Logs ---
function showToast(msg, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '❌';
    if (type === 'warning') icon = '⚠️';

    toast.innerHTML = `<span style="font-size:1.2em">${icon}</span> <span>${msg}</span>`;
    els.toastContainer.appendChild(toast);

    setTimeout(() => {
        toast.style.animation = 'slideIn 0.3s reverse';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// Global Log Function
function log(type, msg) {
    console.log(`[${type}] ${msg}`);

    // 1. Add to DOM
    if (els.logTerminal) {
        const t = new Date().toLocaleTimeString();
        const div = document.createElement('div');
        div.className = `log-line ${type}`;
        div.textContent = `[${t}] ${msg}`;
        els.logTerminal.appendChild(div);
        els.logTerminal.scrollTop = els.logTerminal.scrollHeight;
    }

    // 2. Save to Storage
    const logs = JSON.parse(localStorage.getItem('iqac_logs') || '[]');
    logs.push({ type, msg, time: new Date().toLocaleTimeString() });
    if (logs.length > 200) logs.shift();
    localStorage.setItem('iqac_logs', JSON.stringify(logs));
}

function loadHistory() {
    const logs = JSON.parse(localStorage.getItem('iqac_logs') || '[]');
    if (!els.logTerminal) return;

    els.logTerminal.innerHTML = '<div class="log-line system">System ready.</div>'; // Reset first

    logs.forEach(l => {
        const div = document.createElement('div');
        div.className = `log-line ${l.type}`;
        div.textContent = `[${l.time}] ${l.msg}`;
        els.logTerminal.appendChild(div);
    });
    els.logTerminal.scrollTop = els.logTerminal.scrollHeight;
}

function setupLogs() {
    if (els.clearLogsBtn) {
        els.clearLogsBtn.addEventListener('click', () => {
            if (els.logTerminal) els.logTerminal.innerHTML = '<div class="log-line system">System Ready. Local Logs Cleared.</div>';
            localStorage.removeItem('iqac_logs');
        });
    }
}

// --- Tab Navigation ---
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

// --- CSV ---
function setupCSV() {
    els.csvFile.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) parseCSV(file);
    });

    els.downloadTmplBtn.addEventListener('click', () => {
        const headers = ['Name', 'Email'];
        const csvContent = headers.join(',') + '\nJohn Doe,john.doe@kiit.ac.in';
        const blob = new Blob([csvContent], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'kiit_mailer_template.csv';
        a.click();
        window.URL.revokeObjectURL(url);
    });
}

function parseCSV(file) {
    els.csvStatus.textContent = 'Parsing...';
    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
            if (results.data && results.data.length > 0) {
                state.csvData = results.data;
                state.csvHeaders = results.meta.fields;
                els.csvStatus.textContent = `Loaded ${results.data.length} recipients. ` +
                    `Columns: ${state.csvHeaders.join(', ')}`;
                updatePlaceholders();
                showToast(`CSV Loaded: ${results.data.length} rows`, 'success');
            } else {
                els.csvStatus.textContent = 'Error: No data found';
                showToast('CSV Parse Error: Empty or invalid file', 'error');
            }
        },
        error: (err) => {
            els.csvStatus.textContent = 'Error';
            showToast(`CSV Error: ${err.message}`, 'error');
        }
    });
}

function updatePlaceholders() {
    els.placeholderToolbar.innerHTML = '';
    state.csvHeaders.forEach(header => {
        const btn = document.createElement('button');
        btn.textContent = `{${header}}`;
        btn.className = 'placeholder-btn';
        btn.onclick = () => {
            const range = state.quill.getSelection(true);
            state.quill.insertText(range.index, `{${header}}`);
        };
        els.placeholderToolbar.appendChild(btn);
    });
}

// --- Attachments & Size Limit ---
function setupAttachments() {
    els.attachInput.addEventListener('change', (e) => {
        const newFiles = Array.from(e.target.files);
        let currentSize = state.attachments.reduce((acc, f) => acc + f.size, 0);
        const validFiles = [];
        let skipped = false;

        newFiles.forEach(f => {
            if (currentSize + f.size > MAX_SIZE_MB * 1024 * 1024) {
                skipped = true;
            } else {
                currentSize += f.size;
                validFiles.push(f);
            }
        });

        if (skipped) {
            showToast(`Some files were skipped. Total size exceeds ${MAX_SIZE_MB} MB.`, 'warning');
        }

        state.attachments = [...state.attachments, ...validFiles];
        renderAttachments();
        els.attachInput.value = '';
    });

    els.clearAttachmentsBtn.addEventListener('click', () => {
        state.attachments = [];
        renderAttachments();
        // Maybe reload default?
        // loadDefaultAttachment(); 
        // User asked to clear, so better to clear ALL or restore default? 
        // "Clear" usually means empty. 
    });
}


function renderAttachments() {
    els.attachList.innerHTML = '';

    // Toggle Clear Button Visibility
    if (state.attachments.length > 0) {
        els.clearAttachmentsBtn.style.display = 'inline-block';
    } else {
        els.clearAttachmentsBtn.style.display = 'none';
    }

    let totalSize = 0;

    state.attachments.forEach((file, index) => {
        totalSize += file.size;
        const item = document.createElement('div');
        item.style.cssText = 'background: #333; padding: 4px 8px; border-radius: 4px; display: inline-flex; align-items: center; gap: 8px; margin: 4px; font-size: 0.85rem;';

        const name = document.createElement('span');
        name.textContent = `${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;

        const remove = document.createElement('button');
        remove.textContent = '×';
        remove.style.cssText = 'background: none; border: none; color: #ff5555; cursor: pointer; font-weight: bold;';
        remove.onclick = () => {
            state.attachments.splice(index, 1);
            renderAttachments();
        };

        item.appendChild(name);
        item.appendChild(remove);
        els.attachList.appendChild(item);
    });

    const totalMB = (totalSize / 1024 / 1024).toFixed(2);
    if (totalSize > 0) {
        const info = document.createElement('div');
        info.style.color = totalSize > MAX_SIZE_MB * 1024 * 1024 ? '#ff5555' : '#aaa';
        info.style.fontSize = '0.8rem';
        info.style.marginTop = '4px';
        info.textContent = `Total: ${totalMB} MB / ${MAX_SIZE_MB} MB`;
        els.attachList.appendChild(info);
    }
}

// --- Sending ---
function setupSending() {
    els.sendBtn.addEventListener('click', async () => {
        // Validation of Settings
        if (!state.credentials || !state.credentials.email || !state.credentials.pass) {
            showToast('Please configure settings strictly first!', 'error');
            els.settingsTabBtn.click();
            return;
        }

        const subject = els.subject.value;
        let finalHtmlPayload = '';

        if (state.editorMode === 'html') {
            finalHtmlPayload = els.htmlEditor.value;
            // Validate basic HTML structure
            if (!finalHtmlPayload.trim()) {
                showToast('Email body is empty', 'warning'); return;
            }
        } else {
            const htmlContent = state.quill.root.innerHTML;
            finalHtmlPayload = wrapRichContent(htmlContent);
        }

        if (!subject) { showToast('Please enter a subject', 'warning'); return; }

        if (!state.csvData) { showToast('Please upload a CSV file', 'warning'); return; }

        const emailKey = Object.keys(state.csvData[0]).find(k => k.toLowerCase() === 'email');
        if (!emailKey) {
            showToast("Error: CSV must have an 'Email' column", 'error');
            return;
        }

        const total = state.csvData.length;
        startProgress(total);

        let successCount = 0;
        for (let i = 0; i < total; i++) {
            const row = state.csvData[i];
            const email = row[emailKey] ? row[emailKey].toString().trim() : '';
            if (!email) continue;

            let customHtml = finalHtmlPayload;
            state.csvHeaders.forEach(h => {
                const reg = new RegExp(`{${h}}`, 'g');
                customHtml = customHtml.replace(reg, row[h] || '');
            });

            const ok = await sendOne(email, subject, customHtml, i + 1, total);
            if (ok) successCount++;
        }

        endProgress();
        showToast(`Batch Finished. Sent ${successCount}/${total}`, 'success');
    });
}

function startProgress(total) {
    state.isSending = true;
    els.overlay.classList.remove('hidden');
    els.sendBtn.disabled = true;
    log('system', `Starting batch of ${total} emails...`);
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
    }, 500);
}

async function sendOne(to, subject, html, index, total) {
    updateProgress(index, total);

    const formData = new FormData();
    formData.append('to', to);
    formData.append('subject', subject);
    formData.append('html', html); // This is the processed HTML with replacements

    // Inject Credentials
    if (state.credentials) {
        formData.append('smtpUser', state.credentials.email);
        formData.append('smtpPass', state.credentials.pass);
        if (state.credentials.displayName) formData.append('displayName', state.credentials.displayName);
        if (state.credentials.replyTo) formData.append('replyTo', state.credentials.replyTo);
    }

    state.attachments.forEach(file => formData.append('file', file));

    try {
        const res = await fetch('/api/send-mail', { method: 'POST', body: formData });

        // Read text first to handle non-JSON errors gracefully
        const rawText = await res.text();

        let data;
        try {
            data = JSON.parse(rawText);
        } catch (e) {
            // JSON Parse failed, likely an HTML error page from Vercel
            console.error('Server returned non-JSON:', rawText);
            throw new Error(`Server Error (Raw): ${rawText.substring(0, 100)}...`);
        }

        if (res.ok && data.success) {
            log('success', `Sent to ${to}`);
            return true;
        } else {
            const errMsg = data.error || data.message || 'Unknown Server Error';
            log('error', `Failed to send to ${to}: ${errMsg}`);
            return false;
        }
    } catch (err) {
        log('error', `Error (${to}): ${err.message}`);
        return false;
    }
}
