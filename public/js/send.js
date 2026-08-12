// Batch send engine: chunks recipients into small JSON requests against the
// rewritten /api/send-mail (see api/send-mail.js), retries transient
// failures, persists a LEAN resume snapshot (raw rows + template, not
// resolved HTML per recipient — the old version's localStorage-quota bug),
// and drives the results console (filter, retry-failed-only, CSV export).
import { els } from './dom.js';
import { state, STORAGE_QUEUE, STORAGE_DAILY_COUNT } from './state.js';
import {
    escapeAttr, substitute, hasValidEmailList, dedupeRecipients, toCSV, downloadFile,
} from './util.js';
import {
    getRecipients, currentRoleMap, getEmailFromRow, canonicalizeRow, resolveCopies, computeValidation, currentHeaders,
} from './recipients.js';
import { buildMessagePayload } from './compose.js';
import { getAttachmentRefs, attachmentsSummary } from './attachments.js';
import { showToast, log, openSettingsDrawer, goToStep } from './ui.js';

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
async function sleepInterruptible(ms) {
    const step = 100;
    let elapsed = 0;
    while (elapsed < ms) {
        if (state.isCancelled) return;
        await sleep(Math.min(step, ms - elapsed));
        elapsed += step;
        if (state.isPaused) { while (state.isPaused && !state.isCancelled) await sleep(200); }
    }
}

function buildQueueItems(rows, payload, roleMap) {
    return rows.map((row, idx) => {
        const email = getEmailFromRow(row, roleMap) || (row.Email || '').trim();
        const canonical = canonicalizeRow(row, roleMap);
        const copies = resolveCopies(row, roleMap, payload.rawCc, payload.rawBcc, email);
        return {
            index: idx + 1,
            email,
            cc: copies.cc,
            bcc: copies.bcc,
            invalidCopies: copies.invalid,
            subject: substitute(payload.subject, canonical),
            html: payload.htmlPayload ? substitute(payload.htmlPayload, canonical) : '',
            text: payload.textPayload ? substitute(payload.textPayload, canonical) : '',
            row,
        };
    });
}

function batchSize() { return (state.config && state.config.batchSize) || 10; }

async function callSendMail(items) {
    try {
        const res = await fetch('/api/send-mail', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                smtpUser: state.credentials.email,
                smtpPass: state.credentials.pass,
                displayName: state.credentials.displayName,
                replyTo: state.credentials.replyTo,
                attachmentRefs: state.sendMeta.attachmentRefs,
                messages: items.map(it => ({ to: it.email, cc: it.cc, bcc: it.bcc, subject: it.subject, html: it.html, text: it.text })),
            }),
        });
        const contentType = res.headers.get('content-type') || '';
        if (!contentType.includes('application/json')) {
            const raw = await res.text();
            throw new Error(`Unexpected server response (HTTP ${res.status}): ${raw.substring(0, 150)}`);
        }
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
        return data.results;
    } catch (err) {
        return items.map(() => ({ success: false, error: err.message, retryable: true }));
    }
}

async function sendBatch(items) {
    let results = await callSendMail(items);
    for (let attempt = 1; attempt <= 2 && !state.isCancelled; attempt++) {
        const retryIdxs = results.map((r, i) => (!r.success && r.retryable) ? i : -1).filter(i => i >= 0);
        if (!retryIdxs.length) break;
        await sleep(attempt === 1 ? 2000 : 6000);
        for (const i of retryIdxs) {
            if (state.isCancelled) break;
            const single = await callSendMail([items[i]]);
            results[i] = single[0];
        }
    }
    return results;
}

// ---------- Public entry points ----------
export function setupSending() {
    els.sendBtn.addEventListener('click', () => startFresh());
    els.sendTestBtn.addEventListener('click', sendTestToSelf);
}

function validateCredentials() {
    if (!state.credentials || !state.credentials.email || !state.credentials.pass) {
        showToast('Please configure settings first', 'error');
        openSettingsDrawer();
        return false;
    }
    return true;
}

async function startFresh() {
    if (state.isSending) return;
    if (!validateCredentials()) return;

    const payload = buildMessagePayload();
    if (!payload.subject) { showToast('Please enter a subject', 'warning'); goToStep('compose'); return; }
    if (!payload.htmlPayload.trim() && !payload.textPayload.trim()) {
        showToast('Email body is empty', 'warning'); goToStep('compose'); return;
    }

    const roleMap = currentRoleMap();
    if (state.recipientSource !== 'single' && !roleMap.email) {
        showToast('Map an Email column before sending', 'error'); goToStep('recipients'); return;
    }

    let recipients = getRecipients();
    if (recipients.length === 0) {
        showToast(state.recipientSource === 'single' ? 'Enter recipient email' : 'No recipients to send to', 'warning');
        goToStep('recipients');
        return;
    }
    recipients = recipients.filter(r => hasValidEmailList(getEmailFromRow(r, roleMap)));
    if (recipients.length === 0) {
        showToast('No valid email addresses found', 'error'); goToStep('recipients'); return;
    }
    if (state.recipientSource !== 'single' && els.dedupToggle.checked) {
        const before = recipients.length;
        const normalized = recipients.map(r => ({ ...r, Email: getEmailFromRow(r, roleMap) }));
        recipients = dedupeRecipients(normalized);
        const dropped = before - recipients.length;
        if (dropped > 0) log('system', `Deduplication dropped ${dropped} duplicate email(s)`);
    }

    state.sendMeta = {
        payload, roleMap,
        attachmentRefs: getAttachmentRefs(),
        displayName: state.credentials.displayName, replyTo: state.credentials.replyTo,
    };
    state.sendQueue = buildQueueItems(recipients, payload, roleMap);
    state.sendResults = [];
    state.sendCursor = 0;

    await runQueue(false);
}

async function resumeSaved() {
    const saved = JSON.parse(localStorage.getItem(STORAGE_QUEUE) || 'null');
    if (!saved || !saved.rows || !saved.rows.length) {
        showToast('No saved queue to resume', 'warning');
        return;
    }
    if (!validateCredentials()) return;
    state.sendMeta = {
        payload: saved.payload, roleMap: saved.roleMap,
        attachmentRefs: saved.attachmentRefs || [],
        displayName: saved.displayName, replyTo: saved.replyTo,
    };
    state.sendQueue = buildQueueItems(saved.rows, saved.payload, saved.roleMap);
    state.sendResults = saved.results || [];
    state.sendCursor = 0;
    await runQueue(true);
}

async function retryFailedOnly() {
    const failedEmails = new Set(state.sendResults.filter(r => !r.ok).map(r => r.email));
    const failedItems = state.sendQueue.filter(item => failedEmails.has(item.email));
    if (!failedItems.length) return;
    state.sendResults = state.sendResults.filter(r => r.ok);
    state.sendQueue = failedItems;
    state.sendCursor = 0;
    await runQueue(false, true);
}

async function runQueue(isResume, isRetry) {
    state.isSending = true;
    state.isPaused = false;
    state.isCancelled = false;

    openSendConsole(state.sendQueue.length, isRetry);
    const delay = Math.max(0, parseInt(els.delayMs.value, 10) || 0);
    log('system', `Starting batch of ${state.sendQueue.length} email(s)${isResume ? ' (resumed)' : isRetry ? ' (retry)' : ''}...`);
    persistQueue();

    const bSize = batchSize();
    while (state.sendCursor < state.sendQueue.length) {
        if (state.isCancelled) break;
        if (state.isPaused) { await sleep(200); continue; }

        const batchItems = state.sendQueue.slice(state.sendCursor, state.sendCursor + bSize);
        batchItems.forEach(appendPendingResult);
        const results = await sendBatch(batchItems);
        batchItems.forEach((item, i) => finalizeResult(item, results[i] || { success: false, error: 'No response' }));
        state.sendCursor += batchItems.length;
        updateConsoleProgress();
        persistQueue();

        if (delay > 0 && state.sendCursor < state.sendQueue.length && !state.isCancelled) {
            await sleepInterruptible(delay);
        }
    }

    finishSendConsole();
    if (!state.isCancelled) localStorage.removeItem(STORAGE_QUEUE);
    state.isSending = false;
}

function persistQueue() {
    if (!state.isSending && state.sendQueue.length === 0) return;
    const remaining = state.sendQueue.slice(state.sendCursor);
    if (remaining.length === 0) { localStorage.removeItem(STORAGE_QUEUE); return; }
    const snapshot = {
        savedAt: new Date().toISOString(),
        payload: state.sendMeta.payload,
        roleMap: state.sendMeta.roleMap,
        attachmentRefs: state.sendMeta.attachmentRefs,
        displayName: state.sendMeta.displayName,
        replyTo: state.sendMeta.replyTo,
        rows: remaining.map(item => item.row),
        results: state.sendResults,
    };
    try {
        localStorage.setItem(STORAGE_QUEUE, JSON.stringify(snapshot));
    } catch (e) {
        showToast('Could not save resume checkpoint (browser storage full) — pause/resume across a refresh may not work for this batch.', 'warning');
    }
}

export function checkPendingQueue() {
    const saved = JSON.parse(localStorage.getItem(STORAGE_QUEUE) || 'null');
    if (!saved || !saved.rows || saved.rows.length === 0) {
        els.resumeBanner.classList.add('hidden');
        return;
    }
    const when = saved.savedAt ? new Date(saved.savedAt).toLocaleString() : 'unknown';
    els.resumeMeta.textContent = `${saved.rows.length} unsent recipient(s) from ${when}`;
    els.resumeBanner.classList.remove('hidden');
    els.resumeBtn.onclick = resumeSaved;
    els.discardResumeBtn.onclick = () => {
        if (!confirm('Discard the saved unfinished batch?')) return;
        localStorage.removeItem(STORAGE_QUEUE);
        els.resumeBanner.classList.add('hidden');
        showToast('Saved queue discarded', 'success');
    };
}

// ---------- Send Test to Me ----------
async function sendTestToSelf() {
    if (state.isSending) return;
    if (!validateCredentials()) return;
    const payload = buildMessagePayload();
    if (!payload.subject) { showToast('Please enter a subject', 'warning'); return; }
    if (!payload.htmlPayload.trim() && !payload.textPayload.trim()) {
        showToast('Email body is empty', 'warning'); return;
    }

    const roleMap = currentRoleMap();
    const recips = getRecipients().filter(r => Object.values(r).some(v => (v || '').toString().trim()));
    let row = recips[0];
    if (!row || !getEmailFromRow(row, roleMap)) {
        row = { Name: 'Test User', Email: state.credentials.email };
    }
    const canonical = { ...canonicalizeRow(row, roleMap), Email: state.credentials.email };

    state.sendMeta = {
        payload, roleMap, attachmentRefs: getAttachmentRefs(),
        displayName: state.credentials.displayName, replyTo: state.credentials.replyTo,
    };
    const item = {
        email: state.credentials.email, cc: '', bcc: '',
        subject: '[TEST] ' + substitute(payload.subject, canonical),
        html: payload.htmlPayload ? substitute(payload.htmlPayload, canonical) : '',
        text: payload.textPayload ? substitute(payload.textPayload, canonical) : '',
    };

    els.sendTestBtn.disabled = true;
    showToast('Sending test to your inbox…', 'info');
    log('system', `Sending TEST to ${item.email}`);
    const [result] = await callSendMail([item]);
    els.sendTestBtn.disabled = false;
    if (result.success) showToast('Test sent — check your inbox', 'success');
    else showToast(`Test failed: ${result.error}`, 'error');
}

// ---------- Send console UI ----------
export function setupSendConsole() {
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
    els.closeConsoleBtn.addEventListener('click', () => els.overlay.classList.add('hidden'));
    els.retryFailedBtn.addEventListener('click', () => { els.overlay.classList.add('hidden'); retryFailedOnly(); });

    els.resultsFilterSeg.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            els.resultsFilterSeg.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.resultsFilter = btn.dataset.filter;
            applyResultsFilter();
        });
    });
}

function openSendConsole(total, isRetry) {
    if (!isRetry) els.resultsTable.querySelector('tbody').innerHTML = '';
    els.successCount.textContent = String(state.sendResults.filter(r => r.ok).length);
    els.failureCount.textContent = String(state.sendResults.filter(r => !r.ok).length);
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
    els.retryFailedBtn.style.display = 'none';
    els.closeConsoleBtn.style.display = 'none';
    els.overlay.classList.remove('hidden');
    els.sendBtn.disabled = true;
}

function appendPendingResult(item) {
    const tbody = els.resultsTable.querySelector('tbody');
    const tr = document.createElement('tr');
    tr.className = 'pending';
    tr.dataset.email = item.email;
    tr.innerHTML = `<td>${item.index}</td><td>${escapeAttr(item.email)}</td><td class="status-cell">⏳ Sending…</td><td>—</td>`;
    tbody.appendChild(tr);
    tr.scrollIntoView({ block: 'nearest' });
}

function bumpDailyCount() {
    const today = new Date().toISOString().slice(0, 10);
    let rec;
    try { rec = JSON.parse(localStorage.getItem(STORAGE_DAILY_COUNT) || 'null'); } catch (e) { rec = null; }
    if (!rec || rec.date !== today) rec = { date: today, count: 0 };
    rec.count++;
    localStorage.setItem(STORAGE_DAILY_COUNT, JSON.stringify(rec));
}

function finalizeResult(item, result) {
    const ok = !!result.success;
    if (ok) bumpDailyCount();
    state.sendResults.push({ index: item.index, email: item.email, ok, detail: result.messageId || result.error || '' });
    const tbody = els.resultsTable.querySelector('tbody');
    const tr = tbody.querySelector(`tr[data-email="${CSS.escape(item.email)}"]`);
    if (tr) {
        tr.classList.remove('pending');
        tr.classList.add(ok ? 'success' : 'failure');
        tr.children[2].textContent = ok ? '✓ Sent' : '✕ Failed';
        tr.children[3].textContent = (ok ? result.messageId : result.error) || '';
    }
    applyResultsFilter();
}

function applyResultsFilter() {
    const tbody = els.resultsTable.querySelector('tbody');
    tbody.querySelectorAll('tr').forEach(tr => {
        const show = state.resultsFilter === 'all'
            || (state.resultsFilter === 'success' && tr.classList.contains('success'))
            || (state.resultsFilter === 'failure' && tr.classList.contains('failure'));
        tr.classList.toggle('row-hidden', !show);
    });
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
    els.retryFailedBtn.style.display = fail > 0 ? 'inline-block' : 'none';
    els.closeConsoleBtn.style.display = 'inline-block';
    els.sendBtn.disabled = false;
    updateDailyQuotaUI();
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

// ---------- Review step: summary + pre-flight + quota ----------
export function refreshReviewStep() {
    const creds = state.credentials || {};
    els.summaryFrom.textContent = creds.email ? `${creds.displayName ? creds.displayName + ' · ' : ''}${creds.email}` : 'Not configured';

    const roleMap = currentRoleMap();
    const allRecipients = getRecipients();
    const recipients = allRecipients.filter(r => hasValidEmailList(getEmailFromRow(r, roleMap)));
    els.summaryRecipients.textContent = String(recipients.length);
    els.summaryAttachments.textContent = attachmentsSummary();
    els.sendBtn.textContent = recipients.length > 1 ? `Send to ${recipients.length} recipients` : 'Send Email';

    const payload = buildMessagePayload();
    // CC/BCC coverage and missing-placeholder checks only make sense for rows
    // that will actually be sent (a row with an invalid To address is never
    // queued, so its CC/BCC never goes out either).
    const items = buildQueueItems(recipients, payload, roleMap);
    const ccSet = new Set(), bccSet = new Set(), invalidCopies = new Set();
    items.forEach(it => {
        it.cc.split(',').map(s => s.trim()).filter(Boolean).forEach(e => ccSet.add(e.toLowerCase()));
        it.bcc.split(',').map(s => s.trim()).filter(Boolean).forEach(e => bccSet.add(e.toLowerCase()));
        (it.invalidCopies || []).forEach(e => invalidCopies.add(e));
    });
    els.summaryCopies.textContent = (ccSet.size || bccSet.size)
        ? `${ccSet.size} CC · ${bccSet.size} BCC unique`
        : 'None';

    // Validation stats (invalid/duplicate/missing-name counts) must be computed
    // against the FULL list, not the already-filtered-to-valid one — otherwise
    // "Invalid emails" would always read 0 since invalid rows were pre-removed.
    const v = computeValidation(allRecipients, roleMap);
    // Detect {Token}-shaped placeholders in the body that don't match any
    // known column. The same brace syntax can appear in inline CSS (e.g.
    // "{color-scheme: light only;}"), so this is a heuristic, not a parser —
    // plausible column names are short and don't contain CSS punctuation.
    const tokensInBody = new Set();
    const bodyText = (payload.subject || '') + '\n' + (payload.htmlPayload || '') + '\n' + (payload.textPayload || '');
    bodyText.replace(/\{([^{}]+)\}/g, (_m, k) => {
        const trimmed = k.trim();
        if (trimmed.length <= 40 && !/[:;!]/.test(trimmed)) tokensInBody.add(trimmed);
        return _m;
    });
    const headers = currentHeaders();
    const lowerHeaders = headers.map(h => h.toLowerCase());
    const missing = [...tokensInBody].filter(t => !lowerHeaders.includes(t.toLowerCase()));

    const stats = [
        { label: 'Total recipients', value: v.total, kind: v.total > 0 ? 'ok' : 'warn' },
        { label: 'Invalid emails', value: v.invalid.length, kind: v.invalid.length ? 'danger' : 'ok' },
        { label: 'Invalid CC/BCC', value: invalidCopies.size, kind: invalidCopies.size ? 'danger' : 'ok' },
        { label: 'Duplicates', value: v.duplicates.length, kind: v.duplicates.length ? 'warn' : 'ok' },
        { label: 'Missing placeholders', value: missing.length, kind: missing.length ? 'danger' : 'ok' },
    ];
    let html = stats.map(s => `<div class="preflight-stat ${s.kind}"><span class="label">${s.label}</span><span class="value">${s.value}</span></div>`).join('');
    if (v.invalid.length) html += `<div class="preflight-list"><strong>Invalid:</strong> ${v.invalid.slice(0, 10).map(e => `<span class="tag danger">${escapeAttr(e)}</span>`).join('')}</div>`;
    if (invalidCopies.size) html += `<div class="preflight-list"><strong>Invalid CC/BCC:</strong> ${[...invalidCopies].slice(0, 10).map(e => `<span class="tag danger">${escapeAttr(e)}</span>`).join('')}</div>`;
    if (v.duplicates.length) html += `<div class="preflight-list"><strong>Duplicates:</strong> ${v.duplicates.slice(0, 10).map(e => `<span class="tag warn">${escapeAttr(e)}</span>`).join('')}</div>`;
    if (missing.length) html += `<div class="preflight-list"><strong>Body uses placeholders not in columns:</strong> ${missing.map(m => `<span class="tag">{${escapeAttr(m)}}</span>`).join('')}</div>`;
    els.preflightBody.innerHTML = html;

    updateDailyQuotaUI();
}

function updateDailyQuotaUI() {
    const today = new Date().toISOString().slice(0, 10);
    let rec;
    try { rec = JSON.parse(localStorage.getItem(STORAGE_DAILY_COUNT) || 'null'); } catch (e) { rec = null; }
    const count = (rec && rec.date === today) ? rec.count : 0;
    const limits = (state.config && state.config.gmailDailyLimits) || { personal: 500, workspace: 2000 };
    const pct = Math.min(100, Math.round((count / limits.personal) * 100));
    els.quotaFill.style.width = pct + '%';
    els.quotaFill.classList.toggle('warn', count > limits.personal * 0.7 && count <= limits.personal);
    els.quotaFill.classList.toggle('danger', count > limits.personal);
    els.quotaText.textContent = `${count} sent today · personal Gmail cap ${limits.personal}/day, Google Workspace cap ${limits.workspace}/day`;
}

export function setupReviewRefresh() {
    els.refreshPreflightBtn.addEventListener('click', refreshReviewStep);
}
