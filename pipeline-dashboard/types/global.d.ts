export {};

declare global {
  interface Window {
    storage?: {
      get: (key: string) => Promise<{ value: string | null } | null>;
      set: (key: string, value: string) => Promise<void>;
      delete: (key: string) => Promise<void>;
    };
    fs?: {
      readFile: (name: string) => Promise<Uint8Array>;
      writeFile: (file: File) => Promise<void>;
    };
  }
}
