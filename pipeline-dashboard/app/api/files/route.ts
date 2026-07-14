import { NextRequest, NextResponse } from 'next/server';
import { ensureSchema, sql } from '@/lib/db';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

// GET /api/files -> { files: [{ name, mime, size, uploadedAt }] }
export async function GET() {
  try {
    await ensureSchema();
    const { rows } = await sql`
      SELECT name, mime, size, uploaded_at FROM files ORDER BY uploaded_at DESC
    `;
    return NextResponse.json({
      files: rows.map((r) => ({
        name: r.name,
        mime: r.mime,
        size: Number(r.size),
        uploadedAt: r.uploaded_at,
      })),
    });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'db error' }, { status: 500 });
  }
}

// POST /api/files   multipart form-data with a single "file" field.
export async function POST(req: NextRequest) {
  try {
    await ensureSchema();
    const form = await req.formData();
    const file = form.get('file');
    if (!file || typeof file === 'string') {
      return NextResponse.json({ error: 'No file provided' }, { status: 400 });
    }
    const blob = file as File;
    const buf = Buffer.from(await blob.arrayBuffer());
    const b64 = buf.toString('base64');
    const mime = blob.type || 'application/octet-stream';
    const name = blob.name || 'upload.dat';
    await sql`
      INSERT INTO files (name, mime, data_base64, size, uploaded_at)
      VALUES (${name}, ${mime}, ${b64}, ${buf.length}, now())
      ON CONFLICT (name) DO UPDATE
        SET mime = EXCLUDED.mime,
            data_base64 = EXCLUDED.data_base64,
            size = EXCLUDED.size,
            uploaded_at = now()
    `;
    return NextResponse.json({ ok: true, name, size: buf.length });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'upload error' }, { status: 500 });
  }
}
