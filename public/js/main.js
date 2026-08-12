// Bootstrap: wires every module together. Loaded as <script type="module">
// so plain ES imports work with zero build step.
import { initEls, els } from './dom.js';
import { state, STORAGE_SETTINGS } from './state.js';
import { showToast, log, loadHistory, setupLogs, setupStepper, setupDrawers, goToStep } from './ui.js';
import {
    setupRecipientSourceToggle, setupSingleExtras, setupCSV, setupManualGrid, renderManualGrid,
    setupGlobalCopies, onChange as onRecipientsChanged,
} from './recipients.js';
import {
    initQuill, setupEditorMode, setupSubjectHelpers, updateSubjectHelpers, updatePreheaderCounter,
    refreshPlaceholders, onChange as onComposeChanged,
} from './compose.js';
import { setupAttachments, updateAttachMeter, onChange as onAttachmentsChanged } from './attachments.js';
import { setupPreview, refreshPreview } from './preview.js';
import { setupTemplates, populateTemplateSelect, restoreDraft, scheduleDraftSave } from './storage.js';
import { setupSending, setupSendConsole, checkPendingQueue, refreshReviewStep, setupReviewRefresh } from './send.js';

async function fetchConfig() {
    try {
        const res = await fetch('/api/config');
        state.config = await res.json();
        if (state.config.degraded) {
            els.degradedBannerText.textContent = state.config.degradedReason || 'Running in a reduced-capability mode.';
            els.degradedBanner.classList.remove('hidden');
        }
    } catch (e) {
        // Fall back to conservative defaults (attachments.js/send.js already
        // has a local fallback baked in) — the app still works, just without
        // server-confirmed limits until this succeeds.
        console.warn('Could not fetch /api/config', e);
    }
}

// ---------- Settings (credentials drawer) ----------
function loadSettings() {
    const stored = localStorage.getItem(STORAGE_SETTINGS);
    if (!stored) return;
    try {
        state.credentials = JSON.parse(stored);
        els.settingsEmail.value = state.credentials.email || '';
        els.settingsPass.value = state.credentials.pass || '';
        els.settingsDisplayName.value = state.credentials.displayName || '';
        els.settingsReplyTo.value = state.credentials.replyTo || '';
        applyAccountChip();
    } catch (e) {
        console.error('Failed to parse settings', e);
    }
}

function applyAccountChip() {
    const creds = state.credentials;
    if (!creds || !creds.email) return;
    const dn = creds.displayName || creds.email.split('@')[0];
    els.userNameDisplay.textContent = dn;
    els.userAvatar.textContent = dn.charAt(0).toUpperCase();
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

        state.credentials = { email, pass, displayName, replyTo };
        localStorage.setItem(STORAGE_SETTINGS, JSON.stringify(state.credentials));
        applyAccountChip();
        showToast('Settings saved', 'success');
        els.settingsDrawer.classList.add('hidden');
    });
}

// ---------- Combined change dispatcher ----------
function onAnythingChanged() {
    refreshPlaceholders();
    refreshPreview();
    updateAttachMeter();
    scheduleDraftSave();
    if (state.currentStep === 'review') refreshReviewStep();
}

document.addEventListener('DOMContentLoaded', async () => {
    document.getElementById('year').textContent = new Date().getFullYear();

    initEls();
    loadSettings();
    initQuill();
    await fetchConfig();

    setupStepper((step) => {
        refreshPreview();
        if (step === 'review') refreshReviewStep();
    });
    setupDrawers();
    setupSettings();

    setupRecipientSourceToggle();
    setupSingleExtras();
    setupCSV();
    setupManualGrid();
    setupGlobalCopies();
    onRecipientsChanged(onAnythingChanged);

    setupEditorMode();
    setupSubjectHelpers();
    onComposeChanged(onAnythingChanged);

    setupAttachments();
    onAttachmentsChanged(onAnythingChanged);

    setupPreview();
    setupTemplates();
    setupSending();
    setupSendConsole();
    setupReviewRefresh();
    setupLogs();

    checkPendingQueue();
    loadHistory();

    const restored = restoreDraft();
    if (restored) showToast('Restored your last draft', 'info');

    renderManualGrid();
    populateTemplateSelect();
    refreshPlaceholders();
    updateSubjectHelpers();
    updatePreheaderCounter();
    refreshPreview();

    if (!state.credentials || !state.credentials.email || !state.credentials.pass) {
        showToast('Please configure settings first', 'warning');
        setTimeout(() => els.settingsDrawer.classList.remove('hidden'), 400);
    }
});
