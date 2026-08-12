const express = require('express');
const dotenv = require('dotenv');
const path = require('path');
const portfinder = require('portfinder');
const logger = require('./lib/logger');
// const db = require('./lib/storage'); // DB Removed

// Try loading from .env file (supports both local dev and inside pkg)
dotenv.config({ path: path.join(__dirname, '.env') });
// Fallback: If .env is in CWD (useful if user wants to override config without rebuilding exe)
dotenv.config();

const app = express();

// Middleware for JSON and URL-encoded data
// Increased limits to avoid "PayloadTooLargeError"
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Request Logging Middleware
app.use((req, res, next) => {
    logger.info(`${req.method} ${req.url}`);
    next();
});

// Serve Static Files (Frontend)
app.use(express.static(path.join(__dirname, 'public')));

// API Routes
const sendMailHandler = require('./api/send-mail');
const configHandler = require('./api/config');
const uploadTokenHandler = require('./api/upload-token');
const cleanupHandler = require('./api/cleanup');
const { uploadHandler, serveHandler } = require('./api/upload-local');

function wrap(handler, label) {
    return async (req, res) => {
        try {
            await handler(req, res);
        } catch (e) {
            logger.error("%s Handler Error: %s", label, e.message, { stack: e.stack });
            if (!res.headersSent) {
                res.status(500).json({ error: e.message || 'Internal Server Error' });
            }
        }
    };
}

app.post('/api/send-mail', wrap(sendMailHandler, 'Send Mail'));
app.get('/api/config', wrap(configHandler, 'Config'));
app.post('/api/upload-token', wrap(uploadTokenHandler, 'Upload Token'));
app.post('/api/cleanup', wrap(cleanupHandler, 'Cleanup'));

// Local-only attachment storage (no Vercel Blob store on localhost/.exe runs).
// Raw binary body — kept off the global json/urlencoded parsers above.
app.post('/api/upload', express.raw({ limit: '30mb', type: () => true }), wrap(uploadHandler, 'Local Upload'));
app.get('/api/attachment/:id', wrap(serveHandler, 'Local Attachment Serve'));

// Global Error Handler
app.use((err, req, res, next) => {
    logger.error("Unhandled Express Error: %s", err.message, { stack: err.stack });
    res.status(500).json({ error: 'Internal Server Error' });
});

// Uncaught Exception / Rejection Handlers
process.on('uncaughtException', (err) => {
    logger.error('Uncaught Exception: %s', err.message, { stack: err.stack });
    // In a server, we might want to exit, but for a local user app, keeping it alive might be better, 
    // though risky. Let's log and keep running for now unless it's critical.
});

process.on('unhandledRejection', (reason, promise) => {
    logger.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

// Start Server with Port Finder
const BASE_PORT = process.env.PORT || 3000;
portfinder.basePort = BASE_PORT;

portfinder.getPortPromise()
    .then((port) => {
        app.listen(port, () => {
            const separator = '=================================================';
            console.log(`\n${separator}`);
            console.log(`🚀 IQAC Mailer Local Server Running`);
            console.log(`-------------------------------------------------`);
            console.log(`👉 Access URL  : http://localhost:${port}`);
            console.log(`👉 Environment : ${process.env.NODE_ENV || 'development'}`);
            console.log(`👉 Logs Path   : ${path.join(process.cwd(), 'logs')}`);
            console.log(`👉 Data Path   : ${path.join(process.cwd(), 'data')}`);
            console.log(`${separator}\n`);

            logger.info(`Server started on port ${port}`);
        });
    })
    .catch((err) => {
        logger.error("Failed to find a free port: %s", err.message);
        console.error("Could not find a free port to run the server.");
    });
