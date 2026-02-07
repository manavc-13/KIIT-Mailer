
const express = require('express');
const router = express.Router();
const db = require('../lib/storage');
const logger = require('../lib/logger');

// LIST Drafts
router.get('/', async (req, res) => {
    try {
        const drafts = await db.drafts.find({}).sort({ updatedAt: -1 });
        res.json(drafts);
    } catch (e) {
        logger.error('Error fetching drafts: %s', e.message);
        res.status(500).json({ error: 'Failed to fetch drafts' });
    }
});

// SAVE Draft (Upsert)
router.post('/', async (req, res) => {
    try {
        const { _id, ...data } = req.body;
        // If _id is provided, update; otherwise insert.
        if (_id) {
            const numAffected = await db.drafts.update({ _id }, { $set: data }, {});
            if (numAffected === 0) {
                // If not found, insert as new? Or error? Let's treat as simple save.
                // If user deleted it in another tab, we might want to recreate.
                const newDoc = await db.drafts.insert(data);
                return res.json({ success: true, id: newDoc._id, message: 'Draft created' });
            }
            res.json({ success: true, id: _id, message: 'Draft updated' });
        } else {
            const newDoc = await db.drafts.insert(data);
            res.json({ success: true, id: newDoc._id, message: 'Draft created' });
        }
    } catch (e) {
        logger.error('Error saving draft: %s', e.message);
        res.status(500).json({ error: 'Failed to save draft' });
    }
});

// DELETE Draft
router.delete('/:id', async (req, res) => {
    try {
        await db.drafts.remove({ _id: req.params.id }, {});
        res.json({ success: true });
    } catch (e) {
        logger.error('Error deleting draft: %s', e.message);
        res.status(500).json({ error: 'Failed to delete draft' });
    }
});

module.exports = router;
