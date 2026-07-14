import { sql } from '@vercel/postgres';

let schemaReady: Promise<void> | null = null;

/**
 * Create the tables on first use. Safe to call repeatedly — the work runs once
 * per server instance and is idempotent (CREATE TABLE IF NOT EXISTS).
 *
 * kv_store   — the shared key/value store that replaces the Artifact's
 *              window.storage. All logged-in users read/write the same rows,
 *              so uploaded data is retained and shared across the team.
 * files      — the original uploaded source files, kept so they stay
 *              downloadable ("files should also be available").
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
        data_base64 TEXT NOT NULL,
        size INTEGER NOT NULL DEFAULT 0,
        uploaded_at TIMESTAMPTZ NOT NULL DEFAULT now()
      )`;
    })().catch((e) => {
      // Reset so a later request can retry if the first attempt failed.
      schemaReady = null;
      throw e;
    });
  }
  return schemaReady;
}

export { sql };
