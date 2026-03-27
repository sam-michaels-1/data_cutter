import axios from "axios";
import type {
  UploadResponse,
  DetectColumnsResponse,
  GenerateRequest,
  GenerateResponse,
} from "../types/wizard";

const api = axios.create({ baseURL: "/api" });

export async function uploadFile(file: File): Promise<UploadResponse> {
  const form = new FormData();
  form.append("file", file);
  const { data } = await api.post<UploadResponse>("/upload", form);
  return data;
}

export async function detectColumns(
  sessionId: string,
  sheetName: string
): Promise<DetectColumnsResponse> {
  const { data } = await api.post<DetectColumnsResponse>("/detect-columns", {
    session_id: sessionId,
    sheet_name: sheetName,
  });
  return data;
}

export async function generate(
  req: GenerateRequest
): Promise<GenerateResponse> {
  const { data } = await api.post<GenerateResponse>("/generate", req);
  return data;
}

export function getDownloadUrl(downloadId: string): string {
  return `/api/download/${downloadId}`;
}
