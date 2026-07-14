import { NextRequest, NextResponse } from 'next/server';
import { del } from '@vercel/blob';
import { ensureSchema, sql } from '@/lib/db';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

// GET /api/files/:name
//  - Blob-backed file -> redirect to the CDN URL (no size limit)
//  - inline small file -> stream the bytes back
export async function GET(req: NextRequest, { params }: { params: { name: string } }) {
  try {
    await ensureSchema();
    const name = decodeURIComponent(params.name);
    const { rows } = await sql`
      SELECT mime, url, data_base64 FROM files WHERE name = ${name} LIMIT 1
    `;
    const row = rows[0];
    if (!row) return NextResponse.json({ error: 'not found' }, { status: 404 });

    if (row.url) {
      const sep = row.url.includes('?') ? '&' : '?';
      const download = req.nextUrl.searchParams.get('download') === '1';
      return NextResponse.redirect(download ? `${row.url}${sep}download=1` : row.url);
    }

    if (!row.data_base64) return NextResponse.json({ error: 'not found' }, { status: 404 });
    const buf = Buffer.from(row.data_base64, 'base64');
    const download = req.nextUrl.searchParams.get('download') === '1';
    return new NextResponse(new Uint8Array(buf), {
      status: 200,
      headers: {
        'Content-Type': row.mime || 'application/octet-stream',
        'Content-Length': String(buf.length),
        'Content-Disposition': `${download ? 'attachment' : 'inline'}; filename="${name.replace(/"/g, '')}"`,
        'Cache-Control': 'no-store',
      },
    });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}

// DELETE /api/files/:name — remove the DB row and the Blob object if present.
export async function DELETE(_req: NextRequest, { params }: { params: { name: string } }) {
  try {
    await ensureSchema();
    const name = decodeURIComponent(params.name);
    const { rows } = await sql`SELECT url FROM files WHERE name = ${name} LIMIT 1`;
    if (rows[0]?.url) {
      try {
        await del(rows[0].url);
      } catch {
        /* best-effort blob cleanup */
      }
    }
    await sql`DELETE FROM files WHERE name = ${name}`;
    return NextResponse.json({ ok: true });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}
