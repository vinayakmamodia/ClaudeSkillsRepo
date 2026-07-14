import { SignJWT, jwtVerify } from 'jose';

export const SESSION_COOKIE = 'pd_session';
const ALG = 'HS256';
// Sessions last 7 days.
const MAX_AGE_SECONDS = 60 * 60 * 24 * 7;

function getSecret(): Uint8Array {
  const secret = process.env.AUTH_SECRET;
  if (!secret || secret.length < 16) {
    throw new Error(
      'AUTH_SECRET environment variable is missing or too short (min 16 chars). ' +
        'Set it in your hosting provider / .env.local.'
    );
  }
  return new TextEncoder().encode(secret);
}

/**
 * Parse the APP_USERS env var into a username -> password map.
 * Format: "alice:secret1,bob:secret2"
 * Passwords may themselves contain no comma or colon.
 */
export function getUsers(): Record<string, string> {
  const raw = process.env.APP_USERS || '';
  const users: Record<string, string> = {};
  for (const pair of raw.split(',')) {
    const trimmed = pair.trim();
    if (!trimmed) continue;
    const idx = trimmed.indexOf(':');
    if (idx <= 0) continue;
    const user = trimmed.slice(0, idx).trim();
    const pass = trimmed.slice(idx + 1);
    if (user) users[user] = pass;
  }
  return users;
}

// Constant-time-ish string comparison to avoid trivial timing leaks.
function safeEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let out = 0;
  for (let i = 0; i < a.length; i++) out |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return out === 0;
}

export function verifyCredentials(username: string, password: string): boolean {
  const users = getUsers();
  const expected = users[username];
  if (expected == null) return false;
  return safeEqual(expected, password);
}

export async function createToken(username: string): Promise<string> {
  return new SignJWT({ sub: username })
    .setProtectedHeader({ alg: ALG })
    .setIssuedAt()
    .setExpirationTime(`${MAX_AGE_SECONDS}s`)
    .sign(getSecret());
}

export async function verifyToken(token: string | undefined): Promise<string | null> {
  if (!token) return null;
  try {
    const { payload } = await jwtVerify(token, getSecret(), { algorithms: [ALG] });
    return (payload.sub as string) || null;
  } catch {
    return null;
  }
}

export const sessionCookieOptions = {
  httpOnly: true,
  sameSite: 'lax' as const,
  secure: process.env.NODE_ENV === 'production',
  path: '/',
  maxAge: MAX_AGE_SECONDS,
};
