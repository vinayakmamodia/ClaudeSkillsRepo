import { NextRequest, NextResponse } from 'next/server';
import { ensureSchema, sql } from '@/lib/db';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

// GET /api/files/:name -> the raw file bytes (for download / auto-load)
export async function GET(req: NextRequest, { params }: { params: { name: string } }) {
  try {
    await ensureSchema();
    const name = decodeURIComponent(params.name);
    const { rows } = await sql`
      SELECT mime, data_base64 FROM files WHERE name = ${name} LIMIT 1
    `;
    if (!rows[0]) {
      return NextResponse.json({ error: 'not found' }, { status: 404 });
    }
    const buf = Buffer.from(rows[0].data_base64, 'base64');
    const download = req.nextUrl.searchParams.get('download') === '1';
    const bytes = new Uint8Array(buf);
    return new NextResponse(bytes, {
      status: 200,
      headers: {
        'Content-Type': rows[0].mime || 'application/octet-stream',
        'Content-Length': String(buf.length),
        'Content-Disposition': `${download ? 'attachment' : 'inline'}; filename="${name.replace(/"/g, '')}"`,
        'Cache-Control': 'no-store',
      },
    });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}

// DELETE /api/files/:name
export async function DELETE(_req: NextRequest, { params }: { params: { name: string } }) {
  try {
    await ensureSchema();
    await sql`DELETE FROM files WHERE name = ${decodeURIComponent(params.name)}`;
    return NextResponse.json({ ok: true });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}
