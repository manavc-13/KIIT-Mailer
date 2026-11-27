// Production build for Vercel using the Build Output API v3
// (https://vercel.com/docs/build-output-api/v3). Vercel deploys .vercel/output as-is:
//
//   .vercel/output/static/                    the Vite-built SPA (dist/)
//   .vercel/output/functions/api.func/        ONE Node.js function: the whole Hono API,
//                                              bundled to plain JavaScript with rolldown
//   .vercel/output/config.json                 routing + security headers
//
// Bundling the API ourselves means production never depends on Vercel compiling our
// TypeScript, rewriting `.js` import specifiers, or installing native add-ons.
import { rolldown } from 'rolldown';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const out = path.join(root, '.vercel', 'output');
const fnDir = path.join(out, 'functions', 'api.func');

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(fnDir, { recursive: true });

// 1) Static site (vite build has already produced dist/).
const dist = path.join(root, 'dist');
if (!fs.existsSync(path.join(dist, 'index.html'))) throw new Error('dist/ is missing: run `vite build` first.');
fs.cpSync(dist, path.join(out, 'static'), { recursive: true });

// 2) API function bundle.
const bundle = await rolldown({
  input: path.join(root, 'server', 'vercel.ts'),
  platform: 'node',
  // Optional native speed-ups that `ws` try/catch-requires; absent is fine.
  external: ['bufferutil', 'utf-8-validate'],
  logLevel: 'warn',
});
await bundle.write({
  file: path.join(fnDir, 'index.mjs'),
  format: 'esm',
  inlineDynamicImports: true,
  sourcemap: false,
  minify: false,
  // CommonJS dependencies bundled into ESM still call require() for Node built-ins.
  banner: "import { createRequire as __kmCreateRequire } from 'node:module'; const require = __kmCreateRequire(import.meta.url);",
});
await bundle.close();
fs.writeFileSync(path.join(fnDir, 'package.json'), JSON.stringify({ type: 'module' }));
fs.writeFileSync(path.join(fnDir, '.vc-config.json'), JSON.stringify({
  runtime: 'nodejs22.x',
  handler: 'index.mjs',
  launcherType: 'Nodejs',
  shouldAddHelpers: false,
  maxDuration: 60,
  memory: 1024,
}, null, 2));

// 3) Routing and headers.
const csp = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
  "font-src 'self' https://fonts.gstatic.com data:",
  "img-src 'self' data: blob: https:",
  "connect-src 'self' https://*.vercel-storage.com https://vercel.com",
  "frame-src 'self' blob:",
  "frame-ancestors 'none'",
  "base-uri 'self'",
  "form-action 'self'",
  "object-src 'none'",
].join('; ');
const securityHeaders = {
  'Strict-Transport-Security': 'max-age=63072000; includeSubDomains; preload',
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'X-Frame-Options': 'DENY',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
  'Content-Security-Policy': csp,
};
fs.writeFileSync(path.join(out, 'config.json'), JSON.stringify({
  version: 3,
  routes: [
    { src: '/(.*)', headers: securityHeaders, continue: true },
    { src: '/assets/(.*)', headers: { 'Cache-Control': 'public, max-age=31536000, immutable' }, continue: true },
    // One-click unsubscribe links in emails are /u/<token>; the API serves them under /api/u.
    { src: '^/u/(.*)$', dest: '/api' },
    { src: '^/api(?:/.*)?$', dest: '/api' },
    { handle: 'filesystem' },
    // A missing build asset must 404, never fall through to index.html. Otherwise a tab still
    // running an old deployment would receive HTML for an old JS chunk name.
    { src: '^/assets/.*$', status: 404 },
    // SPA: every other path renders index.html; React Router takes over.
    { src: '/(.*)', dest: '/index.html' },
  ],
}, null, 2));

const kb = (f) => (fs.statSync(f).size / 1024).toFixed(0) + ' KB';
console.log(`Vercel output ready: static/ + functions/api.func/index.mjs (${kb(path.join(fnDir, 'index.mjs'))})`);
