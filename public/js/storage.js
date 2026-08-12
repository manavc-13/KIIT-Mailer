// Named templates (save/load/delete/export/import) and debounced draft
// autosave/restore — so an accidental refresh no longer loses an in-progress
// compose (the old app lost everything on reload).
import { els } from './dom.js';
import { state, STORAGE_TEMPLATES, STORAGE_DRAFT, DEFAULT_HTML, DEFAULT_TEXT } from './state.js';
import { applyEditorMode } from './compose.js';
import { showToast } from './ui.js';
import { downloadFile } from './util.js';

const BUILTIN_NAME = 'KIIT Default (built-in)';

function loadTemplates() {
    try { return JSON.parse(localStorage.getItem(STORAGE_TEMPLATES) || '{}'); }
    catch (e) { return {}; }
}
function saveTemplates(templates) {
    localStorage.setItem(STORAGE_TEMPLATES, JSON.stringify(templates));
}

function currentEditorSnapshot() {
    return {
        subject: els.subject.value,
        preheader: els.preheader.value,
        editorMode: state.editorMode,
        html: els.htmlEditor.value,
        text: els.textEditor.value,
        richHtml: state.quill ? state.quill.root.innerHTML : '',
    };
}

function applySnapshot(snap) {
    els.subject.value = snap.subject || '';
    els.preheader.value = snap.preheader || '';
    els.htmlEditor.value = snap.html || '';
    els.textEditor.value = snap.text || '';
    if (state.quill && snap.richHtml) state.quill.root.innerHTML = snap.richHtml;
    applyEditorMode(snap.editorMode || 'html', false);
}

export function populateTemplateSelect() {
    const templates = loadTemplates();
    const names = Object.keys(templates);
    els.templateSelect.innerHTML = [`<option value="${BUILTIN_NAME}">${BUILTIN_NAME}</option>`]
        .concat(names.map(n => `<option value="${n}">${n}</option>`)).join('');
}

export function setupTemplates() {
    populateTemplateSelect();

    els.loadTemplateBtn.addEventListener('click', () => {
        const name = els.templateSelect.value;
        const hasContent = els.htmlEditor.value.trim() || els.textEditor.value.trim() || els.subject.value.trim();
        if (hasContent && !confirm(`Load "${name}"? This will replace your current subject and body.`)) return;

        if (name === BUILTIN_NAME) {
            applySnapshot({ subject: '', preheader: '', editorMode: 'html', html: DEFAULT_HTML, text: DEFAULT_TEXT, richHtml: '' });
            showToast('Loaded built-in starter template', 'success');
            return;
        }
        const templates = loadTemplates();
        if (!templates[name]) { showToast('Template not found', 'error'); return; }
        applySnapshot(templates[name]);
        showToast(`Loaded template "${name}"`, 'success');
    });

    els.saveTemplateBtn.addEventListener('click', () => {
        els.templateNameInput.value = '';
        els.saveTemplateModal.classList.remove('hidden');
        setTimeout(() => els.templateNameInput.focus(), 50);
    });
    els.cancelSaveTemplate.addEventListener('click', () => els.saveTemplateModal.classList.add('hidden'));
    els.confirmSaveTemplate.addEventListener('click', () => {
        const name = els.templateNameInput.value.trim();
        if (!name) { showToast('Enter a template name', 'warning'); return; }
        if (name === BUILTIN_NAME) { showToast('That name is reserved', 'error'); return; }
        const templates = loadTemplates();
        templates[name] = currentEditorSnapshot();
        saveTemplates(templates);
        populateTemplateSelect();
        els.templateSelect.value = name;
        els.saveTemplateModal.classList.add('hidden');
        showToast(`Saved template "${name}"`, 'success');
    });

    els.deleteTemplateBtn.addEventListener('click', () => {
        const name = els.templateSelect.value;
        if (name === BUILTIN_NAME) { showToast('The built-in template cannot be deleted', 'warning'); return; }
        const templates = loadTemplates();
        if (!templates[name]) return;
        if (!confirm(`Delete template "${name}"?`)) return;
        delete templates[name];
        saveTemplates(templates);
        populateTemplateSelect();
        showToast('Template deleted', 'success');
    });

    els.exportTemplatesBtn.addEventListener('click', () => {
        const templates = loadTemplates();
        if (Object.keys(templates).length === 0) { showToast('No saved templates to export', 'warning'); return; }
        downloadFile(JSON.stringify(templates, null, 2), 'kiit_mailer_templates.json', 'application/json');
    });

    els.importTemplatesInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        try {
            const text = await file.text();
            const incoming = JSON.parse(text);
            const templates = { ...loadTemplates(), ...incoming };
            saveTemplates(templates);
            populateTemplateSelect();
            showToast(`Imported ${Object.keys(incoming).length} template(s)`, 'success');
        } catch (err) {
            showToast(`Import failed: ${err.message}`, 'error');
        }
        e.target.value = '';
    });
}

// ---------- Draft autosave ----------
let draftTimer = null;
export function scheduleDraftSave() {
    clearTimeout(draftTimer);
    draftTimer = setTimeout(saveDraftNow, 800);
}

function saveDraftNow() {
    try {
        const draft = {
            ...currentEditorSnapshot(),
            recipientSource: state.recipientSource,
            ccEmail: els.ccEmail.value,
            bccEmail: els.bccEmail.value,
            manualColumns: state.manualColumns,
            manualRows: state.manualRows,
            manualRoleMap: state.manualRoleMap,
            csvRoleMap: state.csvRoleMap,
            singleExtras: state.singleExtras,
            singleEmail: els.singleEmail.value,
            singleName: els.singleName.value,
            savedAt: new Date().toISOString(),
        };
        localStorage.setItem(STORAGE_DRAFT, JSON.stringify(draft));
    } catch (e) { /* localStorage quota — non-fatal, draft just won't persist */ }
}

export function restoreDraft() {
    let draft;
    try { draft = JSON.parse(localStorage.getItem(STORAGE_DRAFT) || 'null'); }
    catch (e) { return false; }
    if (!draft) return false;

    const hasContent = (draft.html || '').trim() || (draft.text || '').trim() || (draft.subject || '').trim();
    if (!hasContent) return false;

    applySnapshot(draft);
    state.recipientSource = draft.recipientSource || 'single';
    els.ccEmail.value = draft.ccEmail || '';
    els.bccEmail.value = draft.bccEmail || '';
    if (Array.isArray(draft.manualColumns)) state.manualColumns = draft.manualColumns;
    if (Array.isArray(draft.manualRows)) state.manualRows = draft.manualRows;
    if (draft.manualRoleMap) state.manualRoleMap = draft.manualRoleMap;
    if (draft.csvRoleMap) state.csvRoleMap = draft.csvRoleMap;
    if (Array.isArray(draft.singleExtras)) state.singleExtras = draft.singleExtras;
    els.singleEmail.value = draft.singleEmail || '';
    els.singleName.value = draft.singleName || '';

    return true;
}
