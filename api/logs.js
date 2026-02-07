
const express = require('express');
const router = express.Router();
const db = require('../lib/storage');
const logger = require('../lib/logger');

// GET Activity Logs (User Visible)
router.get('/activity', async (req, res) => {
    try {
        const logs = await db.activity.find({}).sort({ createdAt: -1 }).limit(100);
        res.json(logs);
    } catch (e) {
        logger.error('Error fetching activity logs: %s', e.message);
        res.status(500).json({ error: 'Failed to fetch logs' });
    }
});

// GET Sent Mails
router.get('/sent-mails', async (req, res) => {
    try {
        // Limit to last 50 for now, or use pagination if needed later
        const mails = await db.emails.find({}).sort({ createdAt: -1 }).limit(50);
        res.json(mails);
    } catch (e) {
        logger.error('Error fetching sent mails: %s', e.message);
        res.status(500).json({ error: 'Failed to fetch sent mails' });
    }
});

module.exports = router;
