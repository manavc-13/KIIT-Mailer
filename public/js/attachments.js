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
import { state } from './state.js';
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

export function setupAttachments() {
    els.attachInput.addEventListener('change', e => {
        addFiles(Array.from(e.target.files));
        els.attachInput.value = '';
    });
    els.attachDrop.addEventListener('dragover', e => { e.preventDefault(); els.attachDrop.classList.add('dragover'); });
    els.attachDrop.addEventListener('dragleave', () => els.attachDrop.classList.remove('dragover'));
    els.attachDrop.addEventListener('drop', e => {
        e.preventDefault();
        els.attachDrop.classList.remove('dragover');
        addFiles(Array.from(e.dataTransfer.files || []));
    });

    window.addEventListener('beforeunload', () => {
        const ids = state.attachments.filter(a => a.ref && a.ref.provider === 'local').map(a => a.ref.id);
        const urls = state.attachments.filter(a => a.ref && a.ref.provider === 'blob').map(a => a.ref.url);
        if ((ids.length || urls.length) && navigator.sendBeacon) {
            const blob = new Blob([JSON.stringify({ ids, urls })], { type: 'application/json' });
            navigator.sendBeacon('/api/cleanup', blob);
        }
    });
}

function addFiles(files) {
    if (!files.length) return;
    const c = cfg();
    const bodyBytes = currentBodyBytes();
    let runningRaw = readyRawBytes();
    const accepted = [];

    for (const file of files) {
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
    try {
        if (provider === 'local') {
            const res = await fetch(`/api/upload?filename=${encodeURIComponent(file.name)}&contentType=${encodeURIComponent(file.type || 'application/octet-stream')}`, {
                method: 'POST',
                headers: { 'Content-Type': file.type || 'application/octet-stream' },
                body: file,
            });
            const data = await res.json();
            if (!res.ok) throw new Error(data.error || 'Upload failed');
            entry.ref = { provider: 'local', id: data.id, filename: file.name, contentType: file.type, size: file.size };
            entry.status = 'ready';
        } else if (provider === 'blob') {
            // Version pinned to match the server's installed @vercel/blob (package.json)
            // — client/server protocol versions must agree for uploads to work.
            const { upload } = await import('https://esm.sh/@vercel/blob@2.8.0/client');
            const blob = await upload(file.name, file, { access: 'private', handleUploadUrl: '/api/upload-token' });
            entry.ref = { provider: 'blob', url: blob.url, filename: file.name, contentType: file.type, size: file.size };
            entry.status = 'ready';
        } else {
            // inline: no network round-trip, base64 travels with the send request.
            const contentBase64 = await fileToBase64(file);
            entry.ref = { provider: 'inline', contentBase64, filename: file.name, contentType: file.type, size: file.size };
            entry.status = 'ready';
        }
    } catch (err) {
        entry.status = 'error';
        entry.errorMsg = err.message;
        showToast(`Failed to attach "${file.name}": ${err.message}`, 'error');
    }
    renderAttachments();
    onAttachmentsChanged();
}

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

async function removeAttachment(idx) {
    const entry = state.attachments[idx];
    if (!entry) return;
    state.attachments.splice(idx, 1);
    renderAttachments();
    onAttachmentsChanged();
    if (entry.ref && (entry.ref.provider === 'local' || entry.ref.provider === 'blob')) {
        try {
            await fetch('/api/cleanup', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(entry.ref.provider === 'local' ? { ids: [entry.ref.id] } : { urls: [entry.ref.url] }),
            });
        } catch (e) { /* best-effort */ }
    }
}

export function clearAttachments() {
    const ids = state.attachments.filter(a => a.ref && a.ref.provider === 'local').map(a => a.ref.id);
    const urls = state.attachments.filter(a => a.ref && a.ref.provider === 'blob').map(a => a.ref.url);
    state.attachments = [];
    renderAttachments();
    onAttachmentsChanged();
    if (ids.length || urls.length) {
        fetch('/api/cleanup', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ids, urls }),
        }).catch(() => {});
    }
}

export function getAttachmentRefs() {
    return state.attachments.filter(a => a.status === 'ready').map(a => a.ref);
}

export function renderAttachments() {
    els.attachList.innerHTML = '';
    state.attachments.forEach((entry, idx) => {
        const item = document.createElement('div');
        item.className = 'attach-chip';
        const statusIcon = entry.status === 'uploading' ? '⏳' : entry.status === 'error' ? '⚠️' : '📎';
        const statusText = entry.status === 'uploading' ? ' (uploading…)' : entry.status === 'error' ? ` (${entry.errorMsg || 'failed'})` : '';
        item.innerHTML = `
            <span>${statusIcon} ${escapeName(entry.localName)} <span class="muted-small">(${formatBytes(entry.size)})${statusText}</span></span>
            <button type="button" class="x" data-i="${idx}">×</button>
        `;
        item.querySelector('.x').addEventListener('click', () => removeAttachment(idx));
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
