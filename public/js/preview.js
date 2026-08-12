// Persistent Gmail-chrome preview rail: shows the currently selected
// recipient's fully-substituted email, with light/forced-dark and
// desktop/mobile toggles carried over from the old modal-based preview.
import { els } from './dom.js';
import { state } from './state.js';
import { substitute, textToHtml, wrapRichContent } from './util.js';
import { getRecipients, currentRoleMap, canonicalizeRow, computeValidation } from './recipients.js';

export function setupPreview() {
    els.previewPrevBtn.addEventListener('click', () => stepRecipient(-1));
    els.previewNextBtn.addEventListener('click', () => stepRecipient(1));

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
}

function stepRecipient(delta) {
    const recipients = getRecipients();
    if (!recipients.length) return;
    state.previewRowIndex = (state.previewRowIndex + delta + recipients.length) % recipients.length;
    renderPreviewBody();
}

export function refreshPreview() {
    const recipients = getRecipients();
    if (state.previewRowIndex >= recipients.length) state.previewRowIndex = 0;
    els.previewRecipientLabel.textContent = recipients.length ? `${state.previewRowIndex + 1} / ${recipients.length}` : '0 / 0';
    els.previewPrevBtn.disabled = recipients.length <= 1;
    els.previewNextBtn.disabled = recipients.length <= 1;

    const roleMap = currentRoleMap();
    if (recipients.length) {
        const v = computeValidation(recipients, roleMap);
        if (v.invalid.length) {
            els.previewWarnLine.style.display = 'block';
            els.previewWarnLine.textContent = `⚠ ${v.invalid.length} invalid email${v.invalid.length > 1 ? 's' : ''} in this list`;
        } else {
            els.previewWarnLine.style.display = 'none';
        }
    } else {
        els.previewWarnLine.style.display = 'none';
    }

    applyPreviewChrome();
    renderPreviewBody();
}

function applyPreviewChrome() {
    els.gmailFrame.classList.toggle('dark', state.previewTheme === 'dark');
    els.gmailFrame.classList.toggle('mobile', state.previewViewport === 'mobile');
}

function renderPreviewBody() {
    const recipients = getRecipients();
    const roleMap = currentRoleMap();
    const rawRow = recipients[state.previewRowIndex] || { Name: 'John Doe', Email: 'john@example.com' };
    const row = canonicalizeRow(rawRow, roleMap);

    const creds = state.credentials || {};
    const fromName = creds.displayName || (creds.email ? creds.email.split('@')[0] : 'KIIT Mailer');
    els.gmSubject.textContent = substitute(els.subject.value || '(no subject)', row);
    els.gmFromName.textContent = fromName;
    els.gmFromEmail.textContent = creds.email ? `<${creds.email}>` : '<sender@example.com>';
    els.gmTo.textContent = row.Email || 'recipient@example.com';
    els.gmAvatar.textContent = (fromName || 'K').charAt(0).toUpperCase();
    els.gmDate.textContent = new Date().toLocaleString(undefined, { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });

    let html = '';
    if (state.editorMode === 'html') html = els.htmlEditor.value;
    else if (state.editorMode === 'rich') html = wrapRichContent(state.quill.root.innerHTML);
    else html = textToHtml(els.textEditor.value);

    html = substitute(html, row);
    const themedHtml = injectColorScheme(html, state.previewTheme);
    els.previewFrame.onload = () => {
        try {
            const doc = els.previewFrame.contentDocument;
            if (!doc || !doc.body) return;
            const measure = () => {
                const h = Math.max(doc.body.scrollHeight, doc.documentElement.scrollHeight);
                els.previewFrame.style.height = Math.min(Math.max(h, 260), 1600) + 'px';
            };
            measure();
            setTimeout(measure, 120);
        } catch (e) { /* cross-origin: keep CSS height */ }
    };
    els.previewFrame.srcdoc = themedHtml;
}

function injectColorScheme(html, theme) {
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
    if (/<\/head>/i.test(out)) out = out.replace(/<\/head>/i, injected + '</head>');
    else if (/<head[^>]*>/i.test(out)) out = out.replace(/<head[^>]*>/i, m => m + injected);
    else out = injected + out;
    return out;
}
