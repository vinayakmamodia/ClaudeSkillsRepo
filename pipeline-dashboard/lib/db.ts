import { createPool, type VercelPool } from '@vercel/postgres';

// Vercel's Neon/Postgres integrations don't all use the same env var name.
// Accept whichever one is present so the app connects regardless of which
// integration created the database.
export const CONNECTION_VAR_CANDIDATES = [
  'POSTGRES_URL',
  'DATABASE_URL',
  'POSTGRES_PRISMA_URL',
  'POSTGRES_URL_NON_POOLING',
  'DATABASE_URL_UNPOOLED',
];

export function foundConnectionVars(): string[] {
  return CONNECTION_VAR_CANDIDATES.filter((n) => !!(process.env[n] || '').trim());
}

function connectionString(): string {
  for (const name of CONNECTION_VAR_CANDIDATES) {
    const v = (process.env[name] || '').trim();
    if (v) return v;
  }
  throw new Error(
    'No Postgres connection string found. Add the Neon (Postgres) integration in ' +
      'Vercel → Storage and connect it to this project, then redeploy.'
  );
}

let pool: VercelPool | null = null;

function getPool(): VercelPool {
  if (!pool) pool = createPool({ connectionString: connectionString() });
  return pool;
}

// Tagged-template passthrough so routes can keep using: sql`SELECT ...`
export function sql(strings: TemplateStringsArray, ...values: any[]) {
  return getPool().sql(strings, ...values);
}

let schemaReady: Promise<void> | null = null;

/**
 * Create the tables on first use. Idempotent and cached per instance.
 *
 * kv_store — shared key/value store replacing the Artifact's window.storage.
 *            All logged-in users read/write the same rows, so uploaded data is
 *            retained and shared across the team.
 * files    — original uploaded source files, kept so they stay downloadable.
 */
export function ensureSchema(): Promise<void> {
  if (!schemaReady) {
    schemaReady = (async () => {
      await sql`CREATE TABLE IF NOT EXISTS kv_store (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
      )`;
      await sql`CREATE TABLE IF NOT EXISTS files (
        name TEXT PRIMARY KEY,
        mime TEXT NOT NULL DEFAULT 'application/octet-stream',
        data_base64 TEXT,
        size INTEGER NOT NULL DEFAULT 0,
        uploaded_at TIMESTAMPTZ NOT NULL DEFAULT now()
      )`;
      // Large files live in Vercel Blob; store their public URL here.
      // These run every boot but are cheap and idempotent, and they migrate
      // a table that was first created before the `url` column existed.
      await sql`ALTER TABLE files ADD COLUMN IF NOT EXISTS url TEXT`;
      await sql`ALTER TABLE files ALTER COLUMN data_base64 DROP NOT NULL`;
    })().catch((e) => {
      schemaReady = null; // allow retry on next request
      throw e;
    });
  }
  return schemaReady;
}
