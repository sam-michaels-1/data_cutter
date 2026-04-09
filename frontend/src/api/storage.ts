/**
 * Session persistence — IndexedDB for file bytes, sessionStorage for JSON state.
 * Survives accidental page refreshes; auto-expires after 1 hour.
 */
import type { WizardState } from "../types/wizard";
import type { EngineConfig } from "../engine/types";

const DB_NAME = "data-cutter";
const DB_VERSION = 1;
const STORE_NAME = "session";
const TTL_MS = 60 * 60 * 1000; // 1 hour

const SS_WIZARD = "dc-wizard-state";
const SS_CONFIG = "dc-engine-config";
const SS_SESSION = "dc-session";
const SS_SAVED_AT = "dc-saved-at";

// ─── IndexedDB ───────────────────────────────────────────────

let dbPromise: Promise<IDBDatabase> | null = null;

function openDB(): Promise<IDBDatabase> {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise<IDBDatabase>((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => {
      dbPromise = null;
      reject(req.error);
    };
  });
  return dbPromise;
}

export async function saveFileToIDB(
  buffer: ArrayBuffer,
  filename: string,
): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put({
      id: "current",
      fileBuffer: buffer,
      filename,
    });
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function loadFileFromIDB(): Promise<{
  fileBuffer: ArrayBuffer;
  filename: string;
} | null> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const req = tx.objectStore(STORE_NAME).get("current");
    req.onsuccess = () => resolve(req.result ?? null);
    req.onerror = () => reject(req.error);
  });
}

export async function clearIDB(): Promise<void> {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readwrite");
      tx.objectStore(STORE_NAME).delete("current");
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    // If DB can't be opened, nothing to clear
  }
}

// ─── sessionStorage ──────────────────────────────────────────

function touch() {
  sessionStorage.setItem(SS_SAVED_AT, String(Date.now()));
}

export function saveWizardState(state: WizardState): void {
  sessionStorage.setItem(SS_WIZARD, JSON.stringify(state));
  touch();
}

export function loadWizardState(): WizardState | null {
  try {
    const raw = sessionStorage.getItem(SS_WIZARD);
    return raw ? (JSON.parse(raw) as WizardState) : null;
  } catch {
    return null;
  }
}

export function saveEngineConfig(config: EngineConfig): void {
  sessionStorage.setItem(SS_CONFIG, JSON.stringify(config));
  touch();
}

export function loadEngineConfig(): EngineConfig | null {
  try {
    const raw = sessionStorage.getItem(SS_CONFIG);
    return raw ? (JSON.parse(raw) as EngineConfig) : null;
  } catch {
    return null;
  }
}

export function isCacheValid(): boolean {
  const ts = sessionStorage.getItem(SS_SAVED_AT);
  if (!ts) return false;
  const age = Date.now() - Number(ts);
  return age >= 0 && age < TTL_MS;
}

export function clearAllStorage(): void {
  sessionStorage.removeItem(SS_WIZARD);
  sessionStorage.removeItem(SS_CONFIG);
  sessionStorage.removeItem(SS_SESSION);
  sessionStorage.removeItem(SS_SAVED_AT);
  clearIDB();
}
