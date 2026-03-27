import { BrowserRouter, Routes, Route, Navigate } from "react-router-dom";
import SessionProvider from "./components/SessionProvider";
import AppShell from "./components/AppShell";
import ImportPage from "./pages/ImportPage";
import DashboardPage from "./pages/DashboardPage";
import CohortPage from "./pages/CohortPage";
import CustomersPage from "./pages/CustomersPage";
import DownloadPage from "./pages/DownloadPage";

export default function App() {
  return (
    <BrowserRouter>
      <SessionProvider>
        <Routes>
          <Route element={<AppShell />}>
            <Route path="/" element={<Navigate to="/import" replace />} />
            <Route path="/import" element={<ImportPage />} />
            <Route path="/dashboard" element={<DashboardPage />} />
            <Route path="/cohort" element={<CohortPage />} />
            <Route path="/customers" element={<CustomersPage />} />
            <Route path="/download" element={<DownloadPage />} />
          </Route>
        </Routes>
      </SessionProvider>
    </BrowserRouter>
  );
}
