import { NextRequest, NextResponse } from 'next/server';
import { createToken, verifyCredentials, SESSION_COOKIE, sessionCookieOptions } from '@/lib/auth';

export const runtime = 'nodejs';

export async function POST(req: NextRequest) {
  let username = '';
  let password = '';
  try {
    const body = await req.json();
    username = String(body.username || '').trim();
    password = String(body.password || '');
  } catch {
    return NextResponse.json({ error: 'Invalid request' }, { status: 400 });
  }

  if (!username || !password) {
    return NextResponse.json({ error: 'Username and password are required' }, { status: 400 });
  }

  if (!verifyCredentials(username, password)) {
    return NextResponse.json({ error: 'Invalid username or password' }, { status: 401 });
  }

  let token: string;
  try {
    token = await createToken(username);
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || 'Server misconfigured' }, { status: 500 });
  }

  const res = NextResponse.json({ ok: true, username });
  res.cookies.set(SESSION_COOKIE, token, sessionCookieOptions);
  return res;
}
