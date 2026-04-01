import { useCallback } from "react";
import { useSession } from "../components/SessionProvider";
import { getDownloadBlob } from "../api/client";

export default function DownloadPage() {
  const { sessionId } = useSession();

  const handleDownload = useCallback(() => {
    const blob = getDownloadBlob();
    if (!blob) return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "data-pack-output.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, []);

  if (!sessionId) {
    return (
      <div className="flex items-center justify-center h-full min-h-[60vh]">
        <div className="text-center text-gray-500">
          <p className="text-lg font-medium">No data imported yet</p>
          <p className="text-sm mt-1">Import a file first to generate the workbook.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-md mx-auto px-4 py-16 text-center">
      <div className="text-5xl mb-4">&#x2705;</div>
      <h2 className="text-xl font-semibold mb-2">
        Workbook Ready
      </h2>
      <p className="text-gray-500 text-sm mb-6">
        Your analysis workbook has been generated and is ready for download.
      </p>
      <button
        onClick={handleDownload}
        className="inline-block bg-teal-600 text-white px-8 py-3 rounded-lg font-medium hover:bg-teal-700 transition"
      >
        Download Excel
      </button>
    </div>
  );
}
