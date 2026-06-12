// Drop-in replacement for the Claude Artifact `window.storage` API,
// backed by the browser's localStorage. Returns the same `{ value }`
// shape the original dashboard expected.
export const storage = {
  async get(key: string): Promise<{ value: string } | null> {
    try {
      const value = localStorage.getItem(key);
      return value == null ? null : { value };
    } catch {
      return null;
    }
  },
  async set(key: string, value: string): Promise<void> {
    try {
      localStorage.setItem(key, value);
    } catch {
      /* quota / private-mode — ignore, dashboard still works in-memory */
    }
  },
};

// Base path for the auto-updated data files. Vite injects BASE_URL from
// vite.config.ts `base`, so this resolves correctly on GitHub Pages.
export const dataUrl = (file: string) =>
  `${import.meta.env.BASE_URL}data/${file}`.replace(/\/{2,}/g, '/');
