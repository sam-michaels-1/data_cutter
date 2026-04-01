import { createContext, useContext, useState, useCallback } from "react";
import type { Workbook } from "exceljs";
import type { EngineConfig } from "../engine/types";

interface SessionContextValue {
  sessionId: string | null;
  setSessionId: (id: string | null) => void;
  workbook: Workbook | null;
  setWorkbook: (wb: Workbook | null) => void;
  config: EngineConfig | null;
  setConfig: (config: EngineConfig | null) => void;
  downloadUrl: string | null;
  setDownloadUrl: (url: string | null) => void;
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
      }}
    >
      {children}
    </SessionContext.Provider>
  );
}
