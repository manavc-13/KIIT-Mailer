// Vercel function entry. scripts/build-vercel.mjs bundles this file (and everything it
// imports) into one plain-JavaScript ESM file at .vercel/output/functions/api.func/index.mjs,
// so production never depends on Vercel compiling TypeScript or resolving our imports.
import { getRequestListener } from '@hono/node-server';
import { handleRequest } from './app.js';

export default getRequestListener(handleRequest);
