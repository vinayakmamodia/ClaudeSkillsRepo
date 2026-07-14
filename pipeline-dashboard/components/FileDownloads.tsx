'use client';

import { useCallback, useEffect, useState } from 'react';
import { Download, FileDown, Trash2 } from 'lucide-react';
import { FILES_UPDATED_EVENT } from '@/lib/browserShims';

type FileMeta = { name: string; mime: string; size: number; uploadedAt: string };

function fmtSize(n: number) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

export default function FileDownloads() {
  const [files, setFiles] = useState<FileMeta[]>([]);

  const load = useCallback(async () => {
    try {
      const res = await fetch('/api/files', { cache: 'no-store' });
      if (!res.ok) return;
      const data = await res.json();
      setFiles(data.files || []);
    } catch {
      /* ignore */
    }
  }, []);

  useEffect(() => {
    load();
    const handler = () => load();
    window.addEventListener(FILES_UPDATED_EVENT, handler);
    return () => window.removeEventListener(FILES_UPDATED_EVENT, handler);
  }, [load]);

  const remove = async (name: string) => {
    if (!confirm(`Delete "${name}"? This removes the downloadable original file.`)) return;
    await fetch(`/api/files/${encodeURIComponent(name)}`, { method: 'DELETE' });
    load();
  };

  if (files.length === 0) return null;

  return (
    <div className="mt-5 bg-white border border-slate-200 rounded-lg shadow-sm">
      <div className="px-4 py-2.5 border-b border-slate-200 flex items-center gap-2">
        <FileDown className="w-4 h-4 text-slate-500" />
        <span className="text-sm font-semibold text-slate-800">Uploaded source files</span>
        <span className="text-xs text-slate-500">— original files, available to download</span>
      </div>
      <ul className="divide-y divide-slate-100">
        {files.map((f) => (
          <li key={f.name} className="px-4 py-2.5 flex items-center justify-between gap-3">
            <div className="min-w-0">
              <div className="text-sm text-slate-800 truncate" title={f.name}>{f.name}</div>
              <div className="text-[11px] text-slate-500">
                {fmtSize(f.size)} • uploaded {new Date(f.uploadedAt).toLocaleString()}
              </div>
            </div>
            <div className="flex items-center gap-2 flex-shrink-0">
              <a
                href={`/api/files/${encodeURIComponent(f.name)}?download=1`}
                className="flex items-center gap-1.5 text-xs font-medium text-sky-700 hover:text-sky-900 border border-sky-200 hover:border-sky-300 rounded-lg px-2.5 py-1.5 transition"
              >
                <Download className="w-3.5 h-3.5" />
                Download
              </a>
              <button
                onClick={() => remove(f.name)}
                title="Delete file"
                className="text-slate-400 hover:text-rose-600 p-1.5 rounded-lg hover:bg-rose-50 transition"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            </div>
          </li>
        ))}
      </ul>
    </div>
  );
}
