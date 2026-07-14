import { NextRequest, NextResponse } from 'next/server';
import { ensureSchema, sql } from '@/lib/db';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

// GET /api/files -> { files: [{ name, mime, size, uploadedAt, downloadUrl }] }
export async function GET() {
  try {
    await ensureSchema();
    const { rows } = await sql`
      SELECT name, mime, size, uploaded_at, url FROM files ORDER BY uploaded_at DESC
    `;
    return NextResponse.json({
      files: rows.map((r) => ({
        name: r.name,
        mime: r.mime,
        size: Number(r.size),
        uploadedAt: r.uploaded_at,
        // Blob-backed files download straight from the CDN; legacy small files
        // are served back through our route.
        downloadUrl: r.url
          ? `${r.url}?download=1`
          : `/api/files/${encodeURIComponent(r.name)}?download=1`,
      })),
    });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}

// POST /api/files
//  - JSON body { name, url, size, mime }  -> record a Vercel Blob upload (large files)
//  - multipart form-data with "file"      -> store small files inline (fallback)
export async function POST(req: NextRequest) {
  try {
    await ensureSchema();
    const contentType = req.headers.get('content-type') || '';

    if (contentType.includes('application/json')) {
      const { name, url, size, mime } = await req.json();
      if (!name || !url) {
        return NextResponse.json({ error: 'name and url are required' }, { status: 400 });
      }
      await sql`
        INSERT INTO files (name, mime, url, data_base64, size, uploaded_at)
        VALUES (${name}, ${mime || 'application/octet-stream'}, ${url}, NULL, ${Number(size) || 0}, now())
        ON CONFLICT (name) DO UPDATE
          SET mime = EXCLUDED.mime,
              url = EXCLUDED.url,
              data_base64 = NULL,
              size = EXCLUDED.size,
              uploaded_at = now()
      `;
      return NextResponse.json({ ok: true, name });
    }

    // Fallback: small files sent through the function as multipart form-data.
    const form = await req.formData();
    const file = form.get('file');
    if (!file || typeof file === 'string') {
      return NextResponse.json({ error: 'No file provided' }, { status: 400 });
    }
    const blob = file as File;
    const buf = Buffer.from(await blob.arrayBuffer());
    const b64 = buf.toString('base64');
    await sql`
      INSERT INTO files (name, mime, url, data_base64, size, uploaded_at)
      VALUES (${blob.name || 'upload.dat'}, ${blob.type || 'application/octet-stream'}, NULL, ${b64}, ${buf.length}, now())
      ON CONFLICT (name) DO UPDATE
        SET mime = EXCLUDED.mime,
            url = NULL,
            data_base64 = EXCLUDED.data_base64,
            size = EXCLUDED.size,
            uploaded_at = now()
    `;
    return NextResponse.json({ ok: true, name: blob.name, size: buf.length });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'upload error' }, { status: 500 });
  }
}
