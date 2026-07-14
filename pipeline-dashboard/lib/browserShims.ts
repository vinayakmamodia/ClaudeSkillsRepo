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
 *
 * Data is gzip-compressed before it's sent so large parsed datasets stay under
 * the hosting platform's ~4.5 MB request/response body limit.
 */

export const FILES_UPDATED_EVENT = 'pd-files-updated';

let installed = false;

const hasCompression =
  typeof (globalThis as any).CompressionStream !== 'undefined' &&
  typeof (globalThis as any).DecompressionStream !== 'undefined';

async function gzip(str: string): Promise<Uint8Array> {
  const cs = new (globalThis as any).CompressionStream('gzip');
  const writer = cs.writable.getWriter();
  writer.write(new TextEncoder().encode(str));
  writer.close();
  const ab = await new Response(cs.readable).arrayBuffer();
  return new Uint8Array(ab);
}

async function gunzip(buf: ArrayBuffer): Promise<string> {
  const ds = new (globalThis as any).DecompressionStream('gzip');
  const writer = ds.writable.getWriter();
  writer.write(new Uint8Array(buf));
  writer.close();
  const ab = await new Response(ds.readable).arrayBuffer();
  return new TextDecoder().decode(ab);
}

export function installShims() {
  if (installed || typeof window === 'undefined') return;
  installed = true;

  (window as any).storage = {
    async get(key: string) {
      const res = await fetch(`/api/data/${encodeURIComponent(key)}`, { cache: 'no-store' });
      if (!res.ok) return null;
      // Compressed payloads come back as gzip bytes with this header.
      if (res.headers.get('x-encoding') === 'gzip') {
        const buf = await res.arrayBuffer();
        try {
          return { value: await gunzip(buf) };
        } catch {
          return null;
        }
      }
      // Dashboard expects an object shaped like { value: <string|null> }.
      return res.json();
    },
    async set(key: string, value: string) {
      const url = `/api/data/${encodeURIComponent(key)}`;
      if (hasCompression) {
        try {
          const bytes = await gzip(value);
          const res = await fetch(url, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/octet-stream', 'x-encoding': 'gzip' },
            body: bytes,
          });
          if (res.ok) return;
        } catch {
          /* fall back to plain text below */
        }
      }
      await fetch(url, {
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
    // Large files go straight to Vercel Blob (no 4.5 MB limit); if Blob isn't
    // configured, small files fall back to storing inline through the function.
    async writeFile(file: File) {
      try {
        const { upload } = await import('@vercel/blob/client');
        const blob = await upload(file.name, file, {
          access: 'public',
          handleUploadUrl: '/api/files/upload',
          contentType: file.type || undefined,
        });
        await fetch('/api/files', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: file.name, url: blob.url, size: file.size, mime: file.type }),
        });
        window.dispatchEvent(new Event(FILES_UPDATED_EVENT));
        return;
      } catch {
        /* Blob not configured or upload failed — try the inline fallback below. */
      }
      try {
        // Fallback only works for files under the ~4.5 MB function limit.
        if (file.size > 4 * 1024 * 1024) return;
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
