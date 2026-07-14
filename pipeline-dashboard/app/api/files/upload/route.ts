import { handleUpload, type HandleUploadBody } from '@vercel/blob/client';
import { NextRequest, NextResponse } from 'next/server';
import { SESSION_COOKIE, verifyToken } from '@/lib/auth';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

/**
 * Issues a short-lived token so the browser can upload large files DIRECTLY to
 * Vercel Blob, bypassing the ~4.5 MB serverless function body limit.
 *
 * This route is public (see middleware) because Vercel also calls it
 * server-to-server for the upload-completed webhook, which carries no session
 * cookie. Auth is enforced below in onBeforeGenerateToken: no valid session,
 * no upload token.
 */
export async function POST(req: NextRequest): Promise<NextResponse> {
  const body = (await req.json()) as HandleUploadBody;
  try {
    const json = await handleUpload({
      body,
      request: req,
      onBeforeGenerateToken: async () => {
        const user = await verifyToken(req.cookies.get(SESSION_COOKIE)?.value);
        if (!user) throw new Error('Not signed in');
        return {
          allowedContentTypes: [
            'text/csv',
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/octet-stream',
          ],
          maximumSizeInBytes: 200 * 1024 * 1024, // 200 MB ceiling
          // Unique blob path per upload so re-uploading the same filename never
          // hits an "already exists" error. The previous version's blob is
          // cleaned up when the new metadata is recorded (see POST /api/files).
          addRandomSuffix: true,
        };
      },
      // The browser records metadata after upload() resolves, so this is a no-op.
      onUploadCompleted: async () => {},
    });
    return NextResponse.json(json);
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'upload error' }, { status: 400 });
  }
}
