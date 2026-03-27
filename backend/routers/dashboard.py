"""
Dashboard API routes.

GET /api/dashboard/{session_id}  — compute and return dashboard metrics
"""
from __future__ import annotations

import asyncio
import os
from typing import Optional

from fastapi import APIRouter, HTTPException, Query

from schemas_dashboard import DashboardResponse
from temp_store import session_exists, get_session_dir, get_upload_path
from services.compute import compute_dashboard

router = APIRouter(prefix="/api")


@router.get("/dashboard/{session_id}", response_model=DashboardResponse)
async def dashboard(session_id: str, granularity: Optional[str] = Query(None)):
    """Compute dashboard metrics for a given session."""
    if not session_exists(session_id):
        raise HTTPException(404, "Session not found.")

    session_dir = get_session_dir(session_id)
    config_path = os.path.join(session_dir, "config.json")
    if not os.path.isfile(config_path):
        raise HTTPException(404, "Dashboard not available — generate the workbook first.")

    filepath = get_upload_path(session_id)
    if not os.path.isfile(filepath):
        raise HTTPException(404, "Uploaded file not found.")

    loop = asyncio.get_event_loop()
    try:
        result = await loop.run_in_executor(
            None, compute_dashboard, session_dir, filepath, granularity)
    except Exception as e:
        raise HTTPException(500, f"Dashboard computation failed: {e}")

    return result
