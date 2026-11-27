// Local API server for `npm run dev` (Vite proxies /api and /u here).
import { serve } from '@hono/node-server';
import { handleRequest } from './app.js';

const port = Number(process.env.API_PORT || 8787);
serve({ fetch: handleRequest, port }, () => {
  console.log(`KIIT Mailer API listening on http://localhost:${port}/api`);
});
