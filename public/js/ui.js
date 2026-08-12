// Chrome-level UI: toasts, activity log, drawers, stepper navigation.
import { els } from './dom.js';
import { state, STORAGE_LOGS } from './state.js';

export function showToast(msg, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    const icon = { info: 'ℹ️', success: '✅', error: '❌', warning: '⚠️' }[type] || 'ℹ️';
    toast.innerHTML = `<span style="font-size:1.1em">${icon}</span> <span></span>`;
    toast.lastElementChild.textContent = msg;
    els.toastContainer.appendChild(toast);
    setTimeout(() => {
        toast.style.animation = 'slideIn 0.25s reverse';
        setTimeout(() => toast.remove(), 250);
    }, 4200);
}

export function log(type, msg) {
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

export function loadHistory() {
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

export function setupLogs() {
    if (!els.clearLogsBtn) return;
    els.clearLogsBtn.addEventListener('click', () => {
        els.logTerminal.innerHTML = '<div class="log-line system">Local logs cleared.</div>';
        localStorage.removeItem(STORAGE_LOGS);
    });
}

// ---------- Stepper ----------
let onStepChange = null;
export function setupStepper(onChange) {
    onStepChange = onChange;
    els.stepper.querySelectorAll('.step-pill').forEach(btn => {
        btn.addEventListener('click', () => goToStep(btn.dataset.step));
    });
}

export function goToStep(step) {
    state.currentStep = step;
    els.stepper.querySelectorAll('.step-pill').forEach(b => b.classList.toggle('active', b.dataset.step === step));
    els.stepPanels.forEach(p => p.classList.toggle('active', p.dataset.stepPanel === step));
    if (onStepChange) onStepChange(step);
    window.scrollTo({ top: 0, behavior: 'instant' in window ? 'instant' : 'auto' });
}

// ---------- Drawers ----------
export function setupDrawers() {
    const openDrawer = (drawer) => drawer.classList.remove('hidden');
    const closeDrawer = (drawer) => drawer.classList.add('hidden');

    els.accountChip.addEventListener('click', () => openDrawer(els.settingsDrawer));
    els.closeSettingsBtn.addEventListener('click', () => closeDrawer(els.settingsDrawer));
    els.settingsDrawerBackdrop.addEventListener('click', () => closeDrawer(els.settingsDrawer));

    els.activityBtn.addEventListener('click', () => openDrawer(els.activityDrawer));
    els.closeActivityBtn.addEventListener('click', () => closeDrawer(els.activityDrawer));
    els.activityDrawerBackdrop.addEventListener('click', () => closeDrawer(els.activityDrawer));

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            closeDrawer(els.settingsDrawer);
            closeDrawer(els.activityDrawer);
        }
    });
}

export function openSettingsDrawer() {
    els.settingsDrawer.classList.remove('hidden');
}
