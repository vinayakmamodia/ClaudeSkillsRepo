import { cookies } from 'next/headers';
import { redirect } from 'next/navigation';
import { SESSION_COOKIE, verifyToken } from '@/lib/auth';
import DashboardShell from '@/components/DashboardShell';

export const dynamic = 'force-dynamic';

export default async function Home() {
  const token = cookies().get(SESSION_COOKIE)?.value;
  const user = await verifyToken(token);
  if (!user) redirect('/login');

  return <DashboardShell username={user} />;
}
