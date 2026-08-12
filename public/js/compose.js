// Compose step: subject/preheader helpers, editor mode switching (HTML /
// Rich / Plain), placeholder toolbar, and the message payload builder.
import { els } from './dom.js';
import { state, SPAM_WORDS, DEFAULT_HTML, DEFAULT_TEXT } from './state.js';
import { escapeAttr, wrapRichContent, lockLightMode, injectPreheader } from './util.js';
import { currentHeaders } from './recipients.js';

let onBodyChanged = () => {};
export function onChange(fn) { onBodyChanged = fn; }

export function initQuill() {
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
    state.quill.on('text-change', () => onBodyChanged());

    els.htmlEditor.value = DEFAULT_HTML;
    els.textEditor.value = DEFAULT_TEXT;
    els.htmlEditor.addEventListener('input', () => onBodyChanged());
    els.textEditor.addEventListener('input', () => onBodyChanged());
}

export function setupEditorMode() {
    els.editorSeg.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const newMode = btn.dataset.editor;
            if (newMode === state.editorMode) return;
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

export function applyEditorMode(mode, clearOther) {
    state.editorMode = mode;
    els.editorSeg.querySelectorAll('.seg-btn').forEach(b => b.classList.toggle('active', b.dataset.editor === mode));
    els.htmlEditorContainer.style.display = mode === 'html' ? 'block' : 'none';
    els.richEditorContainer.style.display = mode === 'rich' ? 'block' : 'none';
    els.textEditorContainer.style.display = mode === 'text' ? 'block' : 'none';

    if (clearOther) {
        if (mode === 'html') state.quill.setText('');
        else if (mode === 'rich') els.htmlEditor.value = '';
    }
    onBodyChanged();
}

// ---------- Placeholder toolbar ----------
export function refreshPlaceholders() {
    const headers = currentHeaders();
    els.placeholderToolbar.innerHTML = '';
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
    onBodyChanged();
}

function insertAtCursor(textarea, text) {
    const start = textarea.selectionStart || 0;
    const end = textarea.selectionEnd || 0;
    textarea.value = textarea.value.slice(0, start) + text + textarea.value.slice(end);
    textarea.selectionStart = textarea.selectionEnd = start + text.length;
    textarea.focus();
}

// ---------- Subject / preheader helpers ----------
export function setupSubjectHelpers() {
    els.subject.addEventListener('input', () => { updateSubjectHelpers(); onBodyChanged(); });
    els.preheader.addEventListener('input', () => { updatePreheaderCounter(); onBodyChanged(); });
}

export function updateSubjectHelpers() {
    const v = els.subject.value || '';
    const len = v.length;
    els.subjectCounter.textContent = `${len} chars`;
    els.subjectCounter.classList.toggle('warn', len > 60 && len <= 78);
    els.subjectCounter.classList.toggle('danger', len > 78);

    const lower = v.toLowerCase();
    const hits = SPAM_WORDS.filter(w => lower.includes(w.toLowerCase()));
    if (hits.length > 0) {
        els.subjectWarn.style.display = 'block';
        els.subjectWarn.innerHTML = `⚠️ Possible spam-trigger word(s): ${hits.map(h => `<code>${escapeAttr(h)}</code>`).join(', ')}`;
    } else {
        els.subjectWarn.style.display = 'none';
    }
}

export function updatePreheaderCounter() {
    const v = els.preheader.value || '';
    els.preheaderCounter.textContent = `${v.length} chars`;
    els.preheaderCounter.classList.toggle('warn', v.length > 0 && (v.length < 30 || v.length > 110));
}

// ---------- Message payload ----------
// Build current message payload (subject/html/text) from editor state. CC/BCC
// are resolved per-recipient elsewhere (recipients.js resolveCopies) — this
// only returns the RAW global CC/BCC input strings for that merge step.
export function buildMessagePayload() {
    const subject = els.subject.value.trim();
    const rawCc = els.ccEmail.value;
    const rawBcc = els.bccEmail.value;
    let htmlPayload = '';
    let textPayload = '';
    if (state.editorMode === 'html') {
        htmlPayload = els.htmlEditor.value;
    } else if (state.editorMode === 'rich') {
        htmlPayload = wrapRichContent(state.quill.root.innerHTML);
    } else {
        textPayload = els.textEditor.value || '';
    }
    const preheader = (els.preheader.value || '').trim();
    if (preheader && htmlPayload) {
        htmlPayload = injectPreheader(htmlPayload, preheader);
    }
    if (htmlPayload) {
        htmlPayload = lockLightMode(htmlPayload);
    }
    return { subject, htmlPayload, textPayload, preheader, rawCc, rawBcc };
}
