// Applies pending Drizzle migrations from server/db/migrations to DATABASE_URL.
import { migrate } from 'drizzle-orm/neon-serverless/migrator';
import { db } from '../server/db/client.js';

await migrate(db, { migrationsFolder: 'server/db/migrations' });
console.log('Migrations applied.');
process.exit(0);
