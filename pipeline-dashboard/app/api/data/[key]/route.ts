import { NextRequest, NextResponse } from 'next/server';
import { ensureSchema, sql } from '@/lib/db';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

// Compressed values are stored as this marker + base64(gzip-bytes).
const GZ = 'gz:';

// GET /api/data/:key
// Plain values  -> JSON { value: string | null }
// Compressed    -> raw gzip bytes with header x-encoding: gzip
export async function GET(_req: NextRequest, { params }: { params: { key: string } }) {
  try {
    await ensureSchema();
    const { rows } = await sql`SELECT value FROM kv_store WHERE key = ${params.key} LIMIT 1`;
    const v: string | null = rows[0]?.value ?? null;
    if (typeof v === 'string' && v.startsWith(GZ)) {
      const buf = Buffer.from(v.slice(GZ.length), 'base64');
      return new NextResponse(new Uint8Array(buf), {
        status: 200,
        headers: {
          'Content-Type': 'application/octet-stream',
          'x-encoding': 'gzip',
          'Cache-Control': 'no-store',
        },
      });
    }
    return NextResponse.json({ value: v });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}

// PUT /api/data/:key
// x-encoding: gzip  -> body is gzip bytes (stored compressed)
// otherwise         -> body is a raw string (stored as-is)
export async function PUT(req: NextRequest, { params }: { params: { key: string } }) {
  try {
    await ensureSchema();
    let value: string;
    if (req.headers.get('x-encoding') === 'gzip') {
      const buf = Buffer.from(await req.arrayBuffer());
      value = GZ + buf.toString('base64');
    } else {
      value = await req.text();
    }
    await sql`
      INSERT INTO kv_store (key, value, updated_at)
      VALUES (${params.key}, ${value}, now())
      ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()
    `;
    return NextResponse.json({ ok: true });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}

// DELETE /api/data/:key
export async function DELETE(_req: NextRequest, { params }: { params: { key: string } }) {
  try {
    await ensureSchema();
    await sql`DELETE FROM kv_store WHERE key = ${params.key}`;
    return NextResponse.json({ ok: true });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}
