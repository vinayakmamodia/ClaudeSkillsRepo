/**
 * The dashboard was built as a Claude Artifact and talks to two browser-global
 * APIs that only exist inside Claude: `window.storage` (a key/value store) and
 * `window.fs` (a file store). This module reimplements both against our own
 * backend so the dashboard component works unchanged:
 *
 *   window.storage.get/set/delete  ->  /api/data/:key   (shared Postgres KV)
 *   window.fs.readFile             ->  GET  /api/files/:name
 *   window.fs.writeFile            ->  POST /api/files   (keeps originals downloadable)
 *
 * Because the KV store is a single shared table, any data one user uploads is
 * retained and visible to every other logged-in user.
 */

export const FILES_UPDATED_EVENT = 'pd-files-updated';

let installed = false;

export function installShims() {
  if (installed || typeof window === 'undefined') return;
  installed = true;

  (window as any).storage = {
    async get(key: string) {
      const res = await fetch(`/api/data/${encodeURIComponent(key)}`, { cache: 'no-store' });
      if (!res.ok) return null;
      // Dashboard expects an object shaped like { value: <string|null> }.
      return res.json();
    },
    async set(key: string, value: string) {
      await fetch(`/api/data/${encodeURIComponent(key)}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'text/plain' },
        body: value,
      });
    },
    async delete(key: string) {
      await fetch(`/api/data/${encodeURIComponent(key)}`, { method: 'DELETE' });
    },
  };

  (window as any).fs = {
    // Returns a Uint8Array — XLSX.read(buf, { type: 'array' }) accepts this.
    async readFile(name: string) {
      const res = await fetch(`/api/files/${encodeURIComponent(name)}`, { cache: 'no-store' });
      if (!res.ok) throw new Error(`File not found: ${name}`);
      const buf = await res.arrayBuffer();
      return new Uint8Array(buf);
    },
    // Persists the original uploaded File so it stays downloadable.
    async writeFile(file: File) {
      try {
        const fd = new FormData();
        fd.append('file', file);
        await fetch('/api/files', { method: 'POST', body: fd });
        window.dispatchEvent(new Event(FILES_UPDATED_EVENT));
      } catch {
        /* non-fatal: dashboard still works from parsed data in the KV store */
      }
    },
  };
}
