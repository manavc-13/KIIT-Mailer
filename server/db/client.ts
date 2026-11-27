// Single shared Drizzle client over Neon's serverless Pool (WebSocket), which
// supports interactive transactions and SELECT ... FOR UPDATE SKIP LOCKED
// (needed by the job runner). Works on Vercel Node functions and locally.
import { Pool, neonConfig } from '@neondatabase/serverless';
import { drizzle } from 'drizzle-orm/neon-serverless';
import ws from 'ws';
import * as schema from './schema.js';

neonConfig.webSocketConstructor = ws;

const url = process.env.DATABASE_URL;
if (!url) throw new Error('DATABASE_URL is not set');

// idleTimeoutMillis: drop idle connections before Neon / the platform closes them under us.
const pool = new Pool({ connectionString: url, max: 5, idleTimeoutMillis: 10_000 });
// An idle client whose socket dies (Neon compute suspends, network blip) emits 'error' on
// the pool; without a listener Node treats it as an uncaught exception and the process dies.
pool.on('error', (err: Error) => console.error('[db] idle client error', err.message));
export const db = drizzle(pool, { schema });
export type DB = typeof db;
export { schema };
