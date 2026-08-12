// Recipients step: single/CSV/manual sources, column-role mapping (Email /
// Name / CC / BCC), per-row CC/BCC resolution, and validation.
import { els } from './dom.js';
import { state, ROLE_ALIASES } from './state.js';
import {
    escapeAttr, parseEmailList, normalizeEmailList, invalidEmails, hasValidEmailList,
    uniqueEmails, substitute, parseTabular, toCSV, downloadFile,
} from './util.js';
import { showToast } from './ui.js';

let onRecipientsChanged = () => {};
export function onChange(fn) { onRecipientsChanged = fn; }
function notifyChanged() { onRecipientsChanged(); }

// ---------- Column-role detection ----------
function normalizeHeader(h) {
    return String(h || '').replace(/^﻿/, '').trim().toLowerCase();
}

export function detectRoleMap(headers) {
    const map = { email: null, name: null, cc: null, bcc: null };
    const normalized = headers.map(h => ({ raw: h, norm: normalizeHeader(h) }));
    Object.keys(ROLE_ALIASES).forEach(role => {
        const aliases = ROLE_ALIASES[role];
        const hit = normalized.find(h => aliases.includes(h.norm));
        if (hit) map[role] = hit.raw;
    });
    return map;
}

// Add canonical Email/Name aliases into a row copy so {Email}/{Name} always
// work as placeholders even when the source column is named differently
// (e.g. a CSV with "E-mail Address" mapped to the email role).
export function canonicalizeRow(row, roleMap) {
    const out = { ...row };
    if (roleMap.email && row[roleMap.email] != null) out.Email = row[roleMap.email];
    if (roleMap.name && row[roleMap.name] != null) out.Name = row[roleMap.name];
    return out;
}

export function getEmailFromRow(row, roleMap) {
    const key = roleMap.email || 'Email';
    return (row[key] || '').trim();
}

// Merge row CC/BCC with global CC/BCC, de-duplicated case-insensitively,
// never repeating the To address (and BCC never repeats a CC address).
export function resolveCopies(row, roleMap, globalCcRaw, globalBccRaw, toEmail) {
    const globalCc = substitute(globalCcRaw, row);
    const globalBcc = substitute(globalBccRaw, row);
    const rowCcRaw = roleMap.cc ? (row[roleMap.cc] || '') : '';
    const rowBccRaw = roleMap.bcc ? (row[roleMap.bcc] || '') : '';

    const toSet = new Set(parseEmailList(toEmail).map(e => e.toLowerCase()));
    const allInvalid = [
        ...invalidEmails(rowCcRaw), ...invalidEmails(rowBccRaw),
        ...invalidEmails(globalCc), ...invalidEmails(globalBcc),
    ];

    let cc = uniqueEmails(rowCcRaw, globalCc).filter(e => hasValidEmailList(e) && !toSet.has(e.toLowerCase()));
    const ccSet = new Set(cc.map(e => e.toLowerCase()));
    let bcc = uniqueEmails(rowBccRaw, globalBcc).filter(e => hasValidEmailList(e) && !toSet.has(e.toLowerCase()) && !ccSet.has(e.toLowerCase()));

    return {
        cc: cc.join(', '),
        bcc: bcc.join(', '),
        invalid: [...new Set(allInvalid)],
    };
}

// ---------- Headers / recipients accessors ----------
export function currentHeaders() {
    if (state.recipientSource === 'single') {
        const set = new Set(['Name', 'Email']);
        state.singleExtras.forEach(f => { if (f.key && f.key.trim()) set.add(f.key.trim()); });
        return [...set];
    }
    if (state.recipientSource === 'csv') {
        return state.csvHeaders && state.csvHeaders.length ? state.csvHeaders : [];
    }
    return state.manualColumns.slice();
}

export function currentRoleMap() {
    if (state.recipientSource === 'manual') return state.manualRoleMap;
    if (state.recipientSource === 'csv') return state.csvRoleMap;
    return { email: 'Email', name: 'Name', cc: null, bcc: null };
}

export function getRecipients() {
    if (state.recipientSource === 'single') {
        const row = { Name: els.singleName.value || '', Email: normalizeEmailList(els.singleEmail.value) };
        state.singleExtras.forEach(f => { if (f.key) row[f.key] = f.value; });
        return [row];
    }
    if (state.recipientSource === 'csv') return state.csvData || [];
    return state.manualRows.filter(r => Object.values(r).some(v => (v || '').toString().trim()));
}

// ---------- Validation ----------
export function computeValidation(recipients, roleMap) {
    const emailsRaw = recipients.map(r => getEmailFromRow(r, roleMap)).filter(Boolean);
    const emailsLower = emailsRaw.map(e => e.toLowerCase());
    const unique = new Set(emailsLower);
    const invalid = emailsRaw.filter(e => !hasValidEmailList(e));
    const seen = new Set();
    const dupes = [];
    emailsLower.forEach(e => { if (seen.has(e)) dupes.push(e); else seen.add(e); });
    const missingName = roleMap.name ? recipients.filter(r => !String(r[roleMap.name] || '').trim()).length : recipients.length;

    return {
        total: recipients.length,
        uniqueCount: unique.size,
        invalid: [...new Set(invalid)],
        duplicates: [...new Set(dupes)],
        missingName,
    };
}

function renderValidation(statsEl, detailsEl, recipients, roleMap) {
    if (!recipients.length) {
        statsEl.parentElement.style.display = 'none';
        return;
    }
    const v = computeValidation(recipients, roleMap);
    statsEl.parentElement.style.display = 'block';
    statsEl.innerHTML = [
        { label: 'Recipients', value: v.total, kind: v.total ? 'ok' : 'warn' },
        { label: 'Invalid emails', value: v.invalid.length, kind: v.invalid.length ? 'danger' : 'ok' },
        { label: 'Duplicates', value: v.duplicates.length, kind: v.duplicates.length ? 'warn' : 'ok' },
        { label: 'Missing name', value: v.missingName, kind: v.missingName ? 'warn' : 'ok' },
    ].map(s => `<div class="stat-pill ${s.kind}"><span class="label">${s.label}</span><span class="value">${s.value}</span></div>`).join('');

    let html = '';
    if (v.invalid.length) {
        html += `<div class="validation-list"><strong>Invalid:</strong> ${v.invalid.slice(0, 12).map(e => `<span class="tag danger">${escapeAttr(e)}</span>`).join('')}${v.invalid.length > 12 ? ` …+${v.invalid.length - 12} more` : ''}</div>`;
    }
    if (v.duplicates.length) {
        html += `<div class="validation-list"><strong>Duplicates:</strong> ${v.duplicates.slice(0, 12).map(e => `<span class="tag warn">${escapeAttr(e)}</span>`).join('')}${v.duplicates.length > 12 ? ` …+${v.duplicates.length - 12} more` : ''}</div>`;
    }
    detailsEl.innerHTML = html;
}

// ---------- Mapping panel (shared by CSV + manual grid) ----------
const ROLE_LABELS = { email: 'Email', name: 'Name', cc: 'CC', bcc: 'BCC' };

function renderMappingPanel(cardEl, rowsEl, headers, roleMap, onSet) {
    if (!headers.length) { cardEl.style.display = 'none'; return; }
    cardEl.style.display = 'block';
    rowsEl.innerHTML = Object.keys(ROLE_LABELS).map(role => {
        const options = ['<option value="">(none)</option>']
            .concat(headers.map(h => `<option value="${escapeAttr(h)}" ${roleMap[role] === h ? 'selected' : ''}>${escapeAttr(h)}</option>`));
        const isSet = !!roleMap[role];
        return `
            <div class="mapping-row">
                <span class="mapping-role">${ROLE_LABELS[role]}${role === 'email' ? ' *' : ''}</span>
                <select data-role="${role}">${options.join('')}</select>
                <span class="mapping-check ${isSet ? '' : 'missing'}">${isSet ? '✓' : '—'}</span>
            </div>`;
    }).join('');

    rowsEl.querySelectorAll('select').forEach(sel => {
        sel.addEventListener('change', () => {
            onSet(sel.dataset.role, sel.value || null);
        });
    });
}

function refreshCsvMappingAndValidation() {
    renderMappingPanel(els.mappingCard, els.mappingRows, state.csvHeaders, state.csvRoleMap, (role, value) => {
        state.csvRoleMap[role] = value;
        refreshCsvMappingAndValidation();
        notifyChanged();
    });
    renderValidation(els.validationStats, els.validationDetails, state.csvData || [], state.csvRoleMap);
}

function refreshManualMappingAndValidation() {
    renderMappingPanel(els.mappingCardManual, els.mappingRowsManual, state.manualColumns, state.manualRoleMap, (role, value) => {
        state.manualRoleMap[role] = value;
        refreshManualMappingAndValidation();
        notifyChanged();
    });
    const rows = state.manualRows.filter(r => Object.values(r).some(v => (v || '').toString().trim()));
    renderValidation(els.validationStatsManual, els.validationDetailsManual, rows, state.manualRoleMap);
}

// ---------- Recipient source toggle ----------
export function setupRecipientSourceToggle() {
    els.recipientSourceSeg.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            els.recipientSourceSeg.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.recipientSource = btn.dataset.source;
            els.singleFields.style.display = state.recipientSource === 'single' ? 'block' : 'none';
            els.bulkCsv.style.display = state.recipientSource === 'csv' ? 'block' : 'none';
            els.bulkManual.style.display = state.recipientSource === 'manual' ? 'block' : 'none';
            notifyChanged();
        });
    });
}

// ---------- Single recipient extras ----------
export function setupSingleExtras() {
    els.addSingleFieldBtn.addEventListener('click', () => {
        state.singleExtras.push({ key: '', value: '' });
        renderSingleExtras();
    });
    els.singleEmail.addEventListener('input', notifyChanged);
    els.singleName.addEventListener('input', notifyChanged);
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
            notifyChanged();
        });
    });
    els.singleCustomFields.querySelectorAll('[data-remove]').forEach(btn => {
        btn.addEventListener('click', e => {
            const i = +e.currentTarget.dataset.remove;
            state.singleExtras.splice(i, 1);
            renderSingleExtras();
            notifyChanged();
        });
    });
}

// ---------- CSV ----------
export function setupCSV() {
    els.csvFile.addEventListener('change', e => {
        const file = e.target.files[0];
        if (file) parseCSV(file);
    });

    els.csvDrop.addEventListener('dragover', e => { e.preventDefault(); els.csvDrop.classList.add('dragover'); });
    els.csvDrop.addEventListener('dragleave', () => els.csvDrop.classList.remove('dragover'));
    els.csvDrop.addEventListener('drop', e => {
        e.preventDefault();
        els.csvDrop.classList.remove('dragover');
        const file = e.dataTransfer.files && e.dataTransfer.files[0];
        if (file) parseCSV(file);
    });

    els.downloadTmplBtn.addEventListener('click', () => {
        const csv = 'Name,Email,CC,BCC,Department\n' +
            'Dr. A Rao,a.rao@kiit.ac.in,"hod.cse@kiit.ac.in; dean@kiit.ac.in",audit@kiit.ac.in,CSE\n' +
            'John Doe,john.doe@example.com,,,Admin\n';
        downloadFile(csv, 'kiit_mailer_template.csv', 'text/csv');
    });
}

function parseCSV(file) {
    els.csvStatus.textContent = 'Parsing...';
    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        transformHeader: h => h.replace(/^﻿/, '').trim(),
        complete: results => {
            if (results.data && results.data.length > 0) {
                state.csvData = results.data;
                state.csvHeaders = results.meta.fields;
                state.csvRoleMap = detectRoleMap(state.csvHeaders);
                els.csvStatus.textContent = `Loaded ${results.data.length} recipients · Columns: ${state.csvHeaders.join(', ')}`;
                refreshCsvMappingAndValidation();
                if (!state.csvRoleMap.email) {
                    showToast('Could not auto-detect an Email column — map it manually below.', 'warning');
                } else {
                    showToast(`CSV loaded: ${results.data.length} rows`, 'success');
                }
                notifyChanged();
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
export function setupManualGrid() {
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
    });

    els.clearGridBtn.addEventListener('click', () => {
        if (!confirm('Clear all manual entries?')) return;
        state.manualColumns = ['Name', 'Email', 'CC', 'BCC'];
        state.manualRows = [{ Name: '', Email: '', CC: '', BCC: '' }];
        state.manualRoleMap = { email: 'Email', name: 'Name', cc: 'CC', bcc: 'BCC' };
        state.selectedRows.clear();
        renderManualGrid();
    });

    els.deleteSelectedBtn.addEventListener('click', () => {
        const sel = [...state.selectedRows].sort((a, b) => b - a);
        if (sel.length === 0) return;
        if (!confirm(`Delete ${sel.length} selected row(s)?`)) return;
        sel.forEach(i => state.manualRows.splice(i, 1));
        state.selectedRows.clear();
        if (state.manualRows.length === 0) {
            const empty = {}; state.manualColumns.forEach(c => empty[c] = '');
            state.manualRows.push(empty);
        }
        renderManualGrid();
        showToast(`Deleted ${sel.length} row(s)`, 'success');
    });

    els.pasteImportBtn.addEventListener('click', () => {
        const open = els.pasteImportArea.style.display !== 'none';
        els.pasteImportArea.style.display = open ? 'none' : 'block';
        if (!open) setTimeout(() => els.pasteImportText.focus(), 50);
    });
    els.pasteImportCancel.addEventListener('click', () => {
        els.pasteImportArea.style.display = 'none';
        els.pasteImportText.value = '';
    });
    els.pasteImportApply.addEventListener('click', () => {
        const text = els.pasteImportText.value.trim();
        if (!text) { showToast('Nothing to import', 'warning'); return; }
        try {
            const { columns, rows } = parseTabular(text);
            if (!rows.length) { showToast('No rows detected', 'warning'); return; }
            const append = els.pasteImportAppend.checked;
            if (append) {
                columns.forEach(c => { if (!state.manualColumns.includes(c)) state.manualColumns.push(c); });
                state.manualRows = state.manualRows.filter(r => Object.values(r).some(v => (v || '').toString().trim()));
                rows.forEach(r => {
                    const norm = {}; state.manualColumns.forEach(c => norm[c] = r[c] || '');
                    state.manualRows.push(norm);
                });
            } else {
                state.manualColumns = columns.slice();
                if (!state.manualColumns.includes('Email')) state.manualColumns.push('Email');
                if (!state.manualColumns.includes('Name')) state.manualColumns.unshift('Name');
                state.manualRows = rows.map(r => {
                    const norm = {}; state.manualColumns.forEach(c => norm[c] = r[c] || '');
                    return norm;
                });
            }
            state.manualRoleMap = detectRoleMap(state.manualColumns);
            state.selectedRows.clear();
            els.pasteImportArea.style.display = 'none';
            els.pasteImportText.value = '';
            renderManualGrid();
            showToast(`Imported ${rows.length} row(s)`, 'success');
        } catch (err) {
            showToast(`Import failed: ${err.message}`, 'error');
        }
    });

    els.exportGridBtn.addEventListener('click', () => {
        const rows = state.manualRows.filter(r => Object.values(r).some(v => (v || '').toString().trim()));
        if (rows.length === 0) { showToast('Grid is empty', 'warning'); return; }
        downloadFile(toCSV(state.manualColumns, rows), 'kiit_mailer_recipients.csv', 'text/csv');
        showToast(`Exported ${rows.length} row(s)`, 'success');
    });
}

export function renderManualGrid() {
    const t = els.manualGrid;
    t.innerHTML = '';
    state.selectedRows = new Set([...state.selectedRows].filter(i => i < state.manualRows.length));

    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    const thCheck = document.createElement('th');
    thCheck.className = 'head-check';
    const allSelected = state.manualRows.length > 0 && state.selectedRows.size === state.manualRows.length;
    thCheck.innerHTML = `<input type="checkbox" id="masterCheck" ${allSelected ? 'checked' : ''} title="Select all rows">`;
    trh.appendChild(thCheck);

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

    const tbody = document.createElement('tbody');
    state.manualRows.forEach((row, ri) => {
        const tr = document.createElement('tr');
        if (state.selectedRows.has(ri)) tr.classList.add('selected');

        const checkTd = document.createElement('td');
        checkTd.className = 'row-check';
        checkTd.innerHTML = `<input type="checkbox" class="row-select" data-ri="${ri}" ${state.selectedRows.has(ri) ? 'checked' : ''}>`;
        tr.appendChild(checkTd);

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

    const masterCheck = t.querySelector('#masterCheck');
    if (masterCheck) {
        masterCheck.addEventListener('change', e => {
            state.selectedRows = e.target.checked ? new Set(state.manualRows.map((_, i) => i)) : new Set();
            renderManualGrid();
        });
    }
    t.querySelectorAll('.row-select').forEach(cb => {
        cb.addEventListener('change', e => {
            const ri = +e.target.dataset.ri;
            if (e.target.checked) state.selectedRows.add(ri); else state.selectedRows.delete(ri);
            renderManualGrid();
        });
    });
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
            Object.keys(state.manualRoleMap).forEach(role => {
                if (state.manualRoleMap[role] === oldName) state.manualRoleMap[role] = newName;
            });
            renderManualGrid();
        });
    });
    t.querySelectorAll('[data-delcol]').forEach(btn => {
        btn.addEventListener('click', e => {
            const ci = +e.currentTarget.dataset.delcol;
            const col = state.manualColumns[ci];
            state.manualColumns.splice(ci, 1);
            state.manualRows.forEach(r => delete r[col]);
            Object.keys(state.manualRoleMap).forEach(role => {
                if (state.manualRoleMap[role] === col) state.manualRoleMap[role] = null;
            });
            renderManualGrid();
        });
    });
    t.querySelectorAll('tbody input[type="text"], tbody input[type="email"]').forEach(inp => {
        inp.addEventListener('input', e => {
            const ri = +e.target.dataset.ri;
            const col = e.target.dataset.col;
            if (Number.isNaN(ri) || !col) return;
            state.manualRows[ri][col] = e.target.value;
            updateGridStatus();
            notifyChanged();
        });
        inp.addEventListener('paste', e => {
            const text = (e.clipboardData || window.clipboardData).getData('text');
            if (!text || !/[\t\n]/.test(text)) return;
            e.preventDefault();
            const startRi = +e.target.dataset.ri;
            const startCol = e.target.dataset.col;
            const startCi = state.manualColumns.indexOf(startCol);
            if (startCi < 0) return;
            const lines = text.replace(/\r/g, '').replace(/\n+$/, '').split('\n');
            lines.forEach((line, lineIdx) => {
                const cells = line.split('\t');
                const targetRi = startRi + lineIdx;
                while (state.manualRows.length <= targetRi) {
                    const empty = {}; state.manualColumns.forEach(c => empty[c] = '');
                    state.manualRows.push(empty);
                }
                cells.forEach((val, cellIdx) => {
                    const ci = startCi + cellIdx;
                    if (ci >= state.manualColumns.length) return;
                    const col = state.manualColumns[ci];
                    state.manualRows[targetRi][col] = val;
                });
            });
            renderManualGrid();
            showToast(`Pasted ${lines.length} row(s)`, 'success');
        });
    });
    t.querySelectorAll('[data-delrow]').forEach(btn => {
        btn.addEventListener('click', e => {
            const ri = +e.currentTarget.dataset.delrow;
            state.manualRows.splice(ri, 1);
            state.selectedRows.delete(ri);
            state.selectedRows = new Set([...state.selectedRows].map(i => i > ri ? i - 1 : i));
            if (state.manualRows.length === 0) {
                const empty = {}; state.manualColumns.forEach(c => empty[c] = '');
                state.manualRows.push(empty);
            }
            renderManualGrid();
        });
    });

    updateGridStatus();
    refreshManualMappingAndValidation();
    notifyChanged();
}

function updateGridStatus() {
    const valid = state.manualRows.filter(r => (r.Email || '').trim() !== '').length;
    els.gridStatus.textContent = `${state.manualRows.length} row(s) · ${valid} with email`;
    const sel = state.selectedRows.size;
    if (els.deleteSelectedBtn) {
        els.deleteSelectedBtn.style.display = sel > 0 ? 'inline-block' : 'none';
        if (els.selectedCount) els.selectedCount.textContent = sel;
    }
}

// ---------- Global CC/BCC change wiring ----------
export function setupGlobalCopies() {
    els.ccEmail.addEventListener('input', notifyChanged);
    els.bccEmail.addEventListener('input', notifyChanged);
}
