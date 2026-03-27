import { createContext, useContext, useState } from "react";

interface SessionContextValue {
  sessionId: string | null;
  setSessionId: (id: string | null) => void;
}

const SessionContext = createContext<SessionContextValue>({
  sessionId: null,
  setSessionId: () => {},
});

export function useSession() {
  return useContext(SessionContext);
}

export default function SessionProvider({
  children,
}: {
  children: React.ReactNode;
}) {
  const [sessionId, setSessionId] = useState<string | null>(() => {
    if (typeof window !== "undefined") {
      return sessionStorage.getItem("dc-session");
    }
    return null;
  });

  const handleSet = (id: string | null) => {
    setSessionId(id);
    if (id) {
      sessionStorage.setItem("dc-session", id);
    } else {
      sessionStorage.removeItem("dc-session");
    }
  };

  return (
    <SessionContext.Provider value={{ sessionId, setSessionId: handleSet }}>
      {children}
    </SessionContext.Provider>
  );
}
