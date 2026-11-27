import type { schema } from './db/client.js';

export type UserRow = typeof schema.users.$inferSelect;

/** Hono env for every route: `c.get('user')` is set by loadSession (null when signed out). */
export type AppEnv = {
  Variables: {
    user: UserRow | null;
    sessionTokenHash: string | null;
  };
};
