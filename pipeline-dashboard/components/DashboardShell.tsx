'use client';

import { useRouter } from 'next/navigation';
import { LogOut, User } from 'lucide-react';
import { installShims } from '@/lib/browserShims';
import Dashboard from './Dashboard';
import FileDownloads from './FileDownloads';

// Install the window.storage / window.fs shims at module load (browser only),
// before the Dashboard component mounts and its effects run.
installShims();

export default function DashboardShell({ username }: { username: string }) {
  const router = useRouter();

  const logout = async () => {
    await fetch('/api/logout', { method: 'POST' });
    router.replace('/login');
    router.refresh();
  };

  return (
    <div>
      <header className="bg-white border-b border-slate-200 sticky top-0 z-40">
        <div className="max-w-[1500px] mx-auto px-4 md:px-6 py-2.5 flex items-center justify-between gap-4">
          <div className="flex items-center gap-2 text-sm font-semibold text-slate-800">
            <span className="inline-block w-2 h-2 rounded-full bg-emerald-500" />
            Pipeline Dashboard
          </div>
          <div className="flex items-center gap-3">
            <span className="flex items-center gap-1.5 text-xs text-slate-600">
              <User className="w-3.5 h-3.5" />
              {username}
            </span>
            <button
              onClick={logout}
              className="flex items-center gap-1.5 text-xs font-medium text-slate-600 hover:text-rose-700 border border-slate-300 hover:border-rose-300 rounded-lg px-2.5 py-1.5 transition"
            >
              <LogOut className="w-3.5 h-3.5" />
              Sign out
            </button>
          </div>
        </div>
      </header>

      <Dashboard />

      <div className="max-w-[1500px] mx-auto px-4 md:px-6 pb-10">
        <FileDownloads />
      </div>
    </div>
  );
}
