import { NextRequest, NextResponse } from 'next/server';
import { SESSION_COOKIE, verifyToken } from '@/lib/auth';

export const runtime = 'nodejs';

export async function GET(req: NextRequest) {
  const user = await verifyToken(req.cookies.get(SESSION_COOKIE)?.value);
  if (!user) return NextResponse.json({ user: null }, { status: 401 });
  return NextResponse.json({ user });
}
