
const winston = require('winston');
require('winston-daily-rotate-file');
const path = require('path');
const fs = require('fs');

// Ensure logs directory exists only if NOT on Vercel
const logDir = path.join(process.cwd(), 'logs');
if (process.env.VERCEL !== '1' && !fs.existsSync(logDir)) {
    fs.mkdirSync(logDir);
}

// On Vercel (or when VERCEL env is set), do NOT use file logging as filesystem is read-only
const isVercel = process.env.VERCEL === '1';

const transports = [];
if (!isVercel) {
    const transport = new winston.transports.DailyRotateFile({
        filename: path.join(logDir, 'iqac-mailer-%DATE%.log'),
        datePattern: 'YYYY-MM-DD',
        zippedArchive: true,
        maxSize: '20m',
        maxFiles: '14d'
    });
    transports.push(transport);
}

const logger = winston.createLogger({
    level: 'info',
    format: winston.format.combine(
        winston.format.timestamp({
            format: 'YYYY-MM-DD HH:mm:ss'
        }),
        winston.format.errors({ stack: true }),
        winston.format.splat(),
        winston.format.json()
    ),
    defaultMeta: { service: 'iqac-mailer' },
    transports: transports
});

// Always log to console in Vercel or dev
if (process.env.NODE_ENV !== 'production' || isVercel) {
    logger.add(new winston.transports.Console({
        format: winston.format.combine(
            winston.format.colorize(),
            winston.format.simple()
        )
    }));
}

module.exports = logger;
