import { useCallback, useState } from "react";

interface Props {
  onFileSelect: (file: File) => void;
  isLoading: boolean;
  filename: string | null;
}

export default function FileUpload({ onFileSelect, isLoading, filename }: Props) {
  const [dragOver, setDragOver] = useState(false);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file) onFileSelect(file);
    },
    [onFileSelect]
  );

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) onFileSelect(file);
    },
    [onFileSelect]
  );

  if (filename) {
    return (
      <div className="border-2 border-green-300 bg-green-50 rounded-xl p-6 text-center">
        <div className="text-green-600 text-4xl mb-2">{"\u2705"}</div>
        <p className="text-green-800 font-semibold">{filename}</p>
        <p className="text-green-600 text-sm mt-1">File uploaded successfully</p>
      </div>
    );
  }

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
      className={`border-2 border-dashed rounded-xl p-12 text-center transition cursor-pointer
        ${dragOver ? "border-blue-500 bg-blue-50" : "border-gray-300 hover:border-blue-400 hover:bg-gray-50"}
        ${isLoading ? "opacity-50 pointer-events-none" : ""}
      `}
    >
      <input
        type="file"
        accept=".xlsx,.xlsm"
        onChange={handleChange}
        className="hidden"
        id="file-input"
        disabled={isLoading}
      />
      <label htmlFor="file-input" className="cursor-pointer">
        <div className="text-gray-400 text-5xl mb-4">{"\uD83D\uDCC4"}</div>
        <p className="text-gray-600 font-medium">
          {isLoading ? "Uploading..." : "Drop your Excel file here"}
        </p>
        <p className="text-gray-400 text-sm mt-1">or click to browse (.xlsx, .xlsm)</p>
      </label>
    </div>
  );
}
