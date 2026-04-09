import { createContext, useContext, useState, useCallback, useEffect } from "react";
import ExcelJS from "exceljs";
import type { Workbook } from "exceljs";
import type { EngineConfig } from "../engine/types";
import { setCurrentWorkbook, setCurrentConfig } from "../api/client";
import {
  isCacheValid,
  loadFileFromIDB,
  loadEngineConfig,
  clearAllStorage,
} from "../api/storage";

interface SessionContextValue {
  sessionId: string | null;
  setSessionId: (id: string | null) => void;
  workbook: Workbook | null;
  setWorkbook: (wb: Workbook | null) => void;
  config: EngineConfig | null;
  setConfig: (config: EngineConfig | null) => void;
  downloadUrl: string | null;
  setDownloadUrl: (url: string | null) => void;
  restoring: boolean;
}

const SessionContext = createContext<SessionContextValue>({
  sessionId: null,
  setSessionId: () => {},
  workbook: null,
  setWorkbook: () => {},
  config: null,
  setConfig: () => {},
  downloadUrl: null,
  setDownloadUrl: () => {},
  restoring: true,
});

export function useSession() {
  return useContext(SessionContext);
}

export default function SessionProvider({
  children,
}: {
  children: React.ReactNode;
}) {
  const [sessionId, setSessionIdState] = useState<string | null>(() => {
    if (typeof window !== "undefined") {
      return sessionStorage.getItem("dc-session");
    }
    return null;
  });
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [config, setConfig] = useState<EngineConfig | null>(null);
  const [downloadUrl, setDownloadUrlState] = useState<string | null>(null);
  const [restoring, setRestoring] = useState(true);

  // Restore session from IndexedDB + sessionStorage on mount
  useEffect(() => {
    let cancelled = false;

    async function restore() {
      try {
        if (!isCacheValid()) {
          clearAllStorage();
          return;
        }

        const stored = await loadFileFromIDB();
        if (!stored || cancelled) {
          if (!cancelled) clearAllStorage();
          return;
        }

        // Re-parse workbook from stored bytes
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(stored.fileBuffer);
        if (cancelled) return;

        setCurrentWorkbook(wb);
        setWorkbook(wb);

        // Restore engine config
        const savedConfig = loadEngineConfig();
        if (savedConfig) {
          setCurrentConfig(savedConfig);
          setConfig(savedConfig);
        }
      } catch (err) {
        console.warn("Session restore failed, starting fresh:", err);
        if (!cancelled) {
          clearAllStorage();
          setSessionIdState(null);
        }
      } finally {
        if (!cancelled) setRestoring(false);
      }
    }

    restore();
    return () => { cancelled = true; };
  }, []);

  const handleSetSessionId = useCallback((id: string | null) => {
    setSessionIdState(id);
    if (id) {
      sessionStorage.setItem("dc-session", id);
    } else {
      sessionStorage.removeItem("dc-session");
    }
  }, []);

  const handleSetDownloadUrl = useCallback((url: string | null) => {
    // Revoke previous URL to free memory
    setDownloadUrlState((prev) => {
      if (prev) URL.revokeObjectURL(prev);
      return url;
    });
  }, []);

  return (
    <SessionContext.Provider
      value={{
        sessionId,
        setSessionId: handleSetSessionId,
        workbook,
        setWorkbook,
        config,
        setConfig,
        downloadUrl,
        setDownloadUrl: handleSetDownloadUrl,
        restoring,
      }}
    >
      {restoring ? (
        <div className="flex items-center justify-center h-screen">
          <div className="text-center text-gray-500">
            <div className="animate-spin h-8 w-8 border-2 border-teal-500 border-t-transparent rounded-full mx-auto mb-3" />
            <p className="text-sm">Restoring session...</p>
          </div>
        </div>
      ) : (
        children
      )}
    </SessionContext.Provider>
  );
}
