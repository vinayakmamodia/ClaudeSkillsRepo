import { NextRequest, NextResponse } from 'next/server';
import { ensureSchema, sql } from '@/lib/db';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

// GET /api/data/:key -> { value: string | null }
export async function GET(_req: NextRequest, { params }: { params: { key: string } }) {
  try {
    await ensureSchema();
    const { rows } = await sql`SELECT value FROM kv_store WHERE key = ${params.key} LIMIT 1`;
    return NextResponse.json({ value: rows[0]?.value ?? null });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}

// PUT /api/data/:key   body = raw string to store
export async function PUT(req: NextRequest, { params }: { params: { key: string } }) {
  try {
    await ensureSchema();
    const value = await req.text();
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
