
const DraftsURL = '/api/drafts';

async function fetchDrafts() {
    const list = document.getElementById('draftsList');
    list.innerHTML = '<div class="empty-state">Loading drafts...</div>';

    try {
        const res = await fetch(DraftsURL);
        const drafts = await res.json();

        if (drafts.length === 0) {
            list.innerHTML = '<div class="empty-state">No saved drafts found</div>';
            return;
        }

        renderDrafts(drafts);
    } catch (e) {
        console.error(e);
        list.innerHTML = '<div class="empty-state error">Failed to load drafts</div>';
    }
}

function renderDrafts(drafts) {
    const list = document.getElementById('draftsList');
    list.innerHTML = '';

    drafts.forEach(draft => {
        const div = document.createElement('div');
        div.className = 'list-item';

        const date = new Date(draft.updatedAt || draft.createdAt).toLocaleString();
        const subject = draft.subject || '(No Subject)';
        const to = draft.to || '(No Recipient)'; // Only for single mode usually

        div.innerHTML = `
            <div class="item-info">
                <div class="item-title">${subject}</div>
                <div class="item-meta">
                    To: ${to} <span class="separator">•</span> ${date}
                </div>
            </div>
            <div class="item-actions">
                <button class="icon-btn restore-btn" title="Load Draft">✏️</button>
                <button class="icon-btn delete-btn" title="Delete">🗑️</button>
            </div>
        `;

        div.querySelector('.restore-btn').onclick = () => restoreDraft(draft);
        div.querySelector('.delete-btn').onclick = () => deleteDraft(draft._id);

        list.appendChild(div);
    });
}

async function restoreDraft(draft) {
    if (!confirm('Load this draft? Current unsaved changes in composer will be lost.')) return;

    // Populate fields
    if (draft.mode) {
        // Toggle mode logic
        const radio = document.querySelector(`input[name="mode"][value="${draft.mode}"]`);
        if (radio) {
            radio.checked = true;
            radio.dispatchEvent(new Event('change'));
        }
    }

    document.getElementById('toEmail').value = draft.to || '';
    document.getElementById('subject').value = draft.subject || '';
    if (state.quill) state.quill.root.innerHTML = draft.body || '';
    document.getElementById('displayName').value = draft.displayName || '';
    document.getElementById('replyTo').value = draft.replyTo || '';

    // Store current Draft ID so updating overwrites it
    state.currentDraftId = draft._id;

    // Switch to Compose tab
    const composeBtn = document.querySelector('[data-tab="compose"]');
    if (composeBtn) composeBtn.click();

    showToast('Draft loaded', 'success');
}

async function deleteDraft(id) {
    if (!confirm('Delete this draft permanently?')) return;

    try {
        await fetch(`${DraftsURL}/${id}`, { method: 'DELETE' });
        fetchDrafts(); // Refresh
        showToast('Draft deleted', 'success');
    } catch (e) {
        showToast('Failed to delete draft', 'error');
    }
}

// Listeners
document.getElementById('refreshDraftsBtn').addEventListener('click', fetchDrafts);
// Auto-load once when tab is clicked? handled in app.js listener
