// Attachments: upload-once pipeline (local temp dir / Vercel Blob / inline
// base64 fallback — see api/config.js for how the provider is chosen) plus
// the encoded-size budget guard that replaces the old flat 25 MB raw check.
//
// Gmail's real cap is 25 MB of ENCODED message size (base64 inflates raw
// bytes by ~4/3). We size against that, not raw bytes, so the "payload
// error under 15 MB" class of bug can't recur: the number shown to the user
// is the same one the backend enforces (lib/mailer.js), fetched from
// /api/config rather than hardcoded.
import { els } from './dom.js';
import { state, STORAGE_QUEUE } from './state.js';
import { formatBytes, encodedSize } from './util.js';
import { showToast } from './ui.js';

let onAttachmentsChanged = () => {};
export function onChange(fn) { onAttachmentsChanged = fn; }

function cfg() {
    return state.config || {
        provider: 'local', maxRawAttachmentBytes: 18 * 1024 * 1024,
        maxEncodedMessageBytes: 25 * 1024 * 1024, softWarnBytes: 10 * 1024 * 1024,
        mimeOverheadBytes: 96 * 1024,
    };
}

function currentBodyBytes() {
    let html = '';
    if (state.editorMode === 'html') html = els.htmlEditor.value || '';
    else if (state.editorMode === 'rich' && state.quill) html = state.quill.root.innerHTML || '';
    const text = els.textEditor.value || '';
    return new Blob([html, text]).size;
}

function readyRawBytes() {
    return state.attachments.filter(a => a.status !== 'error').reduce((sum, a) => sum + (a.size || 0), 0);
}

// True only for drags that carry real files (not text/links dragged inside
// the editors), so we never hijack normal text drag-and-drop.
function isFileDrag(e) {
    const types = e.dataTransfer && e.dataTransfer.types;
    return !!types && Array.from(types).includes('Files');
}

export function setupAttachments() {
    // The <input type="file"> is hidden, so the drop zone has to open the
    // picker itself (previously clicking it did nothing at all).
    const openPicker = () => els.attachInput.click();
    els.attachDrop.addEventListener('click', openPicker);
    els.attachDrop.addEventListener('keydown', e => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openPicker(); }
    });
    els.attachInput.addEventListener('change', e => {
        addFiles(Array.from(e.target.files || []));
        els.attachInput.value = '';
    });

    // Page-wide file drag handling. Without this, a file released even a few
    // pixels outside the small drop zone makes the browser navigate away to
    // the file, wiping the whole compose session. On the Compose step the
    // entire window accepts the drop; elsewhere the drop is just swallowed.
    // A depth counter stops the highlight flickering as the cursor crosses
    // child elements (dragenter/dragleave fire for every descendant).
    let dragDepth = 0;
    const setHighlight = on => {
        els.attachDrop.classList.toggle('dragover', on);
        document.body.classList.toggle('file-dragging', on);
    };
    window.addEventListener('dragenter', e => {
        if (!isFileDrag(e)) return;
        dragDepth++;
        if (state.currentStep === 'compose') setHighlight(true);
    });
    window.addEventListener('dragleave', e => {
        if (!isFileDrag(e)) return;
        dragDepth = Math.max(0, dragDepth - 1);
        if (dragDepth === 0) setHighlight(false);
    });
    window.addEventListener('dragover', e => {
        if (!isFileDrag(e)) return;
        const handledByZone = e.defaultPrevented; // e.g. the CSV drop zone
        e.preventDefault();
        if (!handledByZone) e.dataTransfer.dropEffect = state.currentStep === 'compose' ? 'copy' : 'none';
    });
    window.addEventListener('drop', e => {
        dragDepth = 0;
        setHighlight(false);
        if (!isFileDrag(e)) return;
        // Already consumed by a more specific zone (e.g. the CSV uploader).
        if (e.defaultPrevented) return;
        e.preventDefault();
        if (state.currentStep === 'compose') {
            addFiles(Array.from(e.dataTransfer.files || []));
        } else {
            showToast('Go to the Compose step to attach files.', 'info');
        }
    });

    window.addEventListener('beforeunload', () => {
        // Keep anything a saved (resumable) batch still points at - deleting
        // it here made "Resume" fail with "attachment not found" after reload.
        const keep = savedQueueRefKeys();
        const refs = state.attachments.map(a => a.ref).filter(r => r && !keep.has(refKey(r)));
        const ids = refs.filter(r => r.provider === 'local').map(r => r.id);
        const urls = refs.filter(r => r.provider === 'blob').map(r => r.url);
        if ((ids.length || urls.length) && navigator.sendBeacon) {
            const blob = new Blob([JSON.stringify({ ids, urls })], { type: 'application/json' });
            navigator.sendBeacon('/api/cleanup', blob);
        }
    });
}

function refKey(ref) {
    if (ref.provider === 'local') return 'local:' + ref.id;
    if (ref.provider === 'blob') return 'blob:' + ref.url;
    return '';
}

function savedQueueRefKeys() {
    try {
        const saved = JSON.parse(localStorage.getItem(STORAGE_QUEUE) || 'null');
        return new Set(((saved && saved.attachmentRefs) || []).map(refKey).filter(Boolean));
    } catch (e) {
        return new Set();
    }
}

function cleanupRefs(refs) {
    const ids = refs.filter(r => r && r.provider === 'local').map(r => r.id);
    const urls = refs.filter(r => r && r.provider === 'blob').map(r => r.url);
    if (!ids.length && !urls.length) return;
    fetch('/api/cleanup', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids, urls }),
    }).catch(() => { /* best-effort */ });
}

function addFiles(files) {
    if (!files.length) return;
    const c = cfg();
    const bodyBytes = currentBodyBytes();
    let runningRaw = readyRawBytes();
    const accepted = [];

    for (const file of files) {
        if (file.size === 0) {
            showToast(`"${file.name}" is empty (or is a folder) and was skipped.`, 'warning');
            continue;
        }
        if (file.size > c.maxRawAttachmentBytes) {
            showToast(`"${file.name}" (${formatBytes(file.size)}) exceeds the ${formatBytes(c.maxRawAttachmentBytes)} per-file limit${c.degraded ? ' for this deployment' : ''}.`, 'error');
            continue;
        }
        const projectedRaw = runningRaw + file.size;
        const projectedEncoded = encodedSize(projectedRaw) + bodyBytes + c.mimeOverheadBytes;
        if (projectedEncoded > c.maxEncodedMessageBytes) {
            showToast(
                `Adding "${file.name}" would make the message ~${formatBytes(projectedEncoded)} encoded — over Gmail's ${formatBytes(c.maxEncodedMessageBytes)} limit.`,
                'error'
            );
            continue;
        }
        runningRaw = projectedRaw;
        accepted.push(file);
    }

    accepted.forEach(file => {
        const entry = {
            localName: file.name, size: file.size, type: file.type || 'application/octet-stream',
            status: 'uploading', ref: null, errorMsg: null,
        };
        state.attachments.push(entry);
        uploadOne(file, entry);
    });
    renderAttachments();
}

async function uploadOne(file, entry) {
    const provider = cfg().provider;
    const contentType = file.type || 'application/octet-stream';
    try {
        if (provider === 'local') {
            // Body is always sent as octet-stream so the server's JSON/form
            // parsers never touch it (a .json attachment used to be swallowed).
            const res = await fetch(`/api/upload?filename=${encodeURIComponent(file.name)}&contentType=${encodeURIComponent(contentType)}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/octet-stream' },
                body: file,
            });
            const data = await readJson(res);
            if (!res.ok) throw new Error(data.error || `Upload failed (HTTP ${res.status})`);
            entry.ref = { provider: 'local', id: data.id, filename: file.name, contentType, size: file.size };
        } else if (provider === 'blob') {
            let upload;
            try {
                // Version pinned to match the server's installed @vercel/blob (package.json)
                // - client/server protocol versions must agree for uploads to work.
                ({ upload } = await import('https://esm.sh/@vercel/blob@2.8.0/client'));
            } catch (e) {
                throw new Error('could not load the upload library from esm.sh - check the internet connection or whether esm.sh is blocked on this network');
            }
            const blob = await upload(file.name, file, { access: 'private', handleUploadUrl: '/api/upload-token', contentType });
            entry.ref = { provider: 'blob', url: blob.url, filename: file.name, contentType, size: file.size };
        } else {
            // inline: no network round-trip, base64 travels with the send request.
            const contentBase64 = await fileToBase64(file);
            entry.ref = { provider: 'inline', contentBase64, filename: file.name, contentType, size: file.size };
        }
        entry.status = 'ready';
    } catch (err) {
        entry.status = 'error';
        entry.errorMsg = err.message;
        showToast(`Failed to attach "${file.name}": ${err.message}`, 'error');
    }
    // Removed by the user while it was still uploading: don't leave the
    // uploaded copy orphaned on disk / in Blob storage.
    if (!state.attachments.includes(entry)) {
        if (entry.ref) cleanupRefs([entry.ref]);
        return;
    }
    renderAttachments();
    onAttachmentsChanged();
}

async function readJson(res) {
    const text = await res.text();
    try { return JSON.parse(text); }
    catch (e) { return { error: `Unexpected server response (HTTP ${res.status}): ${text.slice(0, 120)}` }; }
}

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

function removeAttachment(entry) {
    const idx = state.attachments.indexOf(entry);
    if (idx < 0) return;
    state.attachments.splice(idx, 1);
    renderAttachments();
    onAttachmentsChanged();
    if (entry.ref && !savedQueueRefKeys().has(refKey(entry.ref))) cleanupRefs([entry.ref]);
}

export function clearAttachments() {
    const refs = state.attachments.map(a => a.ref).filter(Boolean);
    state.attachments = [];
    renderAttachments();
    onAttachmentsChanged();
    cleanupRefs(refs);
}

// Attachments still uploading (or failed) used to be silently dropped from
// the send. Returns a user-facing reason to block the send, or null.
export function attachmentsBlockingSend() {
    const uploading = state.attachments.filter(a => a.status === 'uploading').length;
    if (uploading) {
        return uploading > 1
            ? `${uploading} attachments are still uploading - wait for them to finish.`
            : '1 attachment is still uploading - wait for it to finish.';
    }
    const failed = state.attachments.filter(a => a.status === 'error');
    if (failed.length) {
        return `Remove or re-attach the failed attachment(s): ${failed.map(a => a.localName).join(', ')}`;
    }
    return null;
}

export function getAttachmentRefs() {
    return state.attachments.filter(a => a.status === 'ready').map(a => a.ref);
}

export function renderAttachments() {
    els.attachList.innerHTML = '';
    state.attachments.forEach(entry => {
        const item = document.createElement('div');
        item.className = 'attach-chip';
        const statusIcon = entry.status === 'uploading' ? '⏳' : entry.status === 'error' ? '⚠️' : '📎';
        const statusText = entry.status === 'uploading' ? ' (uploading…)' : entry.status === 'error' ? ` (${entry.errorMsg || 'failed'})` : '';
        item.classList.toggle('error', entry.status === 'error');
        item.innerHTML = `
            <span class="attach-name">${statusIcon} ${escapeName(entry.localName)} <span class="muted-small">(${formatBytes(entry.size)})${escapeName(statusText)}</span></span>
            <button type="button" class="x" title="Remove attachment" aria-label="Remove ${escapeName(entry.localName)}">×</button>
        `;
        item.querySelector('.x').addEventListener('click', () => removeAttachment(entry));
        els.attachList.appendChild(item);
    });
    updateAttachMeter();
}

function escapeName(s) {
    const div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
}

export function updateAttachMeter() {
    const c = cfg();
    const rawBytes = readyRawBytes();
    const bodyBytes = currentBodyBytes();
    const encoded = encodedSize(rawBytes) + bodyBytes + c.mimeOverheadBytes;
    const pct = Math.min(100, Math.round((encoded / c.maxEncodedMessageBytes) * 100));

    if (state.attachments.length === 0) {
        els.attachMeter.style.display = 'none';
        els.attachWarnLine.style.display = 'none';
        return;
    }
    els.attachMeter.style.display = 'block';
    els.attachMeterFill.style.width = pct + '%';
    els.attachMeterFill.classList.toggle('warn', encoded > c.softWarnBytes && encoded <= c.maxEncodedMessageBytes);
    els.attachMeterFill.classList.toggle('danger', encoded > c.maxEncodedMessageBytes);
    els.attachMeterText.textContent =
        `${formatBytes(rawBytes)} raw ≈ ${formatBytes(encoded)} encoded / ${formatBytes(c.maxEncodedMessageBytes)} budget (${pct}%)`;

    if (rawBytes > c.softWarnBytes) {
        els.attachWarnLine.style.display = 'block';
        els.attachWarnLine.className = 'warn-line';
        els.attachWarnLine.textContent = 'Large attachments — some recipient mail servers reject messages this size even though it fits Gmail\'s limit. Consider a shared Drive link instead.';
    } else {
        els.attachWarnLine.style.display = 'none';
    }
}

export function attachmentsSummary() {
    const n = state.attachments.filter(a => a.status !== 'error').length;
    if (!n) return 'None';
    const bytes = readyRawBytes();
    return `${n} file${n > 1 ? 's' : ''} · ${formatBytes(bytes)}`;
}
