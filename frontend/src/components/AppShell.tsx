import { useState } from "react";
import { Outlet } from "react-router-dom";
import Sidebar from "./Sidebar";

export default function AppShell() {
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  return (
    <div className="flex min-h-screen bg-gray-50 text-gray-900">
      {/* Desktop sidebar */}
      <div className="hidden md:block">
        <Sidebar />
      </div>

      {/* Mobile header */}
      <div className="fixed top-0 left-0 right-0 z-40 md:hidden flex items-center gap-2.5 px-4 py-3 bg-gray-900 text-gray-300">
        <button
          onClick={() => setMobileMenuOpen(true)}
          className="p-1 -ml-1 text-gray-400 hover:text-gray-200"
          aria-label="Open menu"
        >
          <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5" />
          </svg>
        </button>
        <img src="/cube-cutter-logo.png" alt="Cube Cutter" className="w-8 h-8 rounded" />
        <span className="text-lg font-bold text-teal-400 tracking-tight">
          Cube Cutter
        </span>
      </div>

      {/* Mobile drawer overlay */}
      {mobileMenuOpen && (
        <div className="fixed inset-0 z-50 md:hidden">
          {/* Backdrop */}
          <div
            className="absolute inset-0 bg-black/50"
            onClick={() => setMobileMenuOpen(false)}
          />
          {/* Drawer */}
          <div className="relative w-64 h-full">
            <Sidebar onClose={() => setMobileMenuOpen(false)} />
          </div>
        </div>
      )}

      <main className="flex-1 overflow-auto pt-14 md:pt-0">
        <Outlet />
      </main>
    </div>
  );
}
