const nodemailer = require('nodemailer');
const Busboy = require('busboy');
const logger = require('../lib/logger');
// const db = require('../lib/storage');

// Vercel config to disable default body parsing for this route
// This allows busboy to consume the stream directly.
module.exports.config = {
    api: {
        bodyParser: false,
        responseLimit: false,
    },
};

module.exports = async (req, res) => {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
    }

    const busboy = Busboy({ headers: req.headers });
    const fields = {};
    const files = [];

    busboy.on('field', (fieldname, val) => {
        fields[fieldname] = val;
    });

    busboy.on('file', (fieldname, file, filename, encoding, mimetype) => {
        const chunks = [];
        file.on('data', data => chunks.push(data));
        file.on('end', () => {
            let actualFilename = filename;
            if (typeof filename === 'object' && filename.filename) {
                actualFilename = filename.filename;
            }

            files.push({
                filename: actualFilename,
                content: Buffer.concat(chunks),
                encoding,
                mimetype
            });
        });
    });

    busboy.on('finish', async () => {
        try {
            // Log attempt
            logger.info('Attempting to send email to: %s', fields.to);

            // Configuration from request
            const smtpUser = fields.smtpUser;
            const smtpPass = fields.smtpPass;

            if (!smtpUser || !smtpPass) {
                throw new Error("Missing SMTP Credentials (User/Pass)");
            }

            const transporter = nodemailer.createTransport({
                service: 'gmail', // Use service: 'gmail' for simplicity with standard gmail accounts
                auth: {
                    user: smtpUser,
                    pass: smtpPass
                }
            });

            const mailOptions = {
                from: `"${fields.displayName || 'EDGEI 2026'}" <${smtpUser}>`,
                to: fields.to,
                replyTo: fields.replyTo || undefined,
                subject: fields.subject,
                html: fields.html, // HTML Body
                attachments: files // Nodemailer handles Buffer content
            };

            const info = await transporter.sendMail(mailOptions);

            logger.info('Email sent successfully. MessageID: %s', info.messageId);

            res.status(200).json({ success: true, messageId: info.messageId });

        } catch (error) {
            logger.error("Mail Error: %s", error.message);
            res.status(500).json({ error: error.message });
        }
    });

    busboy.on('error', (err) => {
        logger.error('Busboy error: %s', err);
        if (!res.headersSent) res.status(500).json({ error: 'Upload failed' });
    });

    req.pipe(busboy);
};
