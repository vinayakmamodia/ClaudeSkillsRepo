import { NextResponse } from 'next/server';
import { ensureSchema, sql, foundConnectionVars } from '@/lib/db';
import { getUsers } from '@/lib/auth';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

/**
 * Safe diagnostic. Requires login (guarded by middleware). Reports ONLY:
 *  - which connection-string env var NAMES are present (never their values)
 *  - how many login accounts APP_USERS parsed to (never the passwords)
 *  - whether the database is reachable, and current row counts
 * No secrets are ever returned.
 */
export async function GET() {
  const connectionVars = foundConnectionVars();
  const result: Record<string, unknown> = {
    connectionStringFound: connectionVars.length > 0,
    connectionVarNames: connectionVars,
    userCount: Object.keys(getUsers()).length,
    authSecretSet: !!(process.env.AUTH_SECRET || '').trim(),
    db: 'not-tested',
  };

  try {
    await ensureSchema();
    const r = await sql`SELECT
      (SELECT count(*) FROM kv_store) AS kv,
      (SELECT count(*) FROM files) AS files`;
    result.db = 'ok';
    result.kvRows = Number(r.rows[0].kv);
    result.fileRows = Number(r.rows[0].files);

    // Non-sensitive detail: which keys/files exist and how big (bytes).
    const kv = await sql`SELECT key, octet_length(value) AS bytes FROM kv_store ORDER BY key`;
    result.kvKeys = kv.rows.map((x) => ({ key: x.key, bytes: Number(x.bytes) }));
    const fl = await sql`SELECT name, size FROM files ORDER BY name`;
    result.files = fl.rows.map((x) => ({ name: x.name, bytes: Number(x.size) }));
  } catch (e: any) {
    result.db = 'error';
    result.dbError = e?.message || String(e);
  }

  return NextResponse.json(result);
}
