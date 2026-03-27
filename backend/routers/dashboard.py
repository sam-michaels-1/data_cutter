"""
Dashboard API routes.

GET /api/dashboard/{session_id}  — compute and return dashboard metrics
"""
from __future__ import annotations

import asyncio
import json
import os
from functools import partial
from typing import Optional

from fastapi import APIRouter, HTTPException, Query

from schemas_dashboard import DashboardResponse
from temp_store import session_exists, get_session_dir, get_upload_path
from services.compute import compute_dashboard

router = APIRouter(prefix="/api")


@router.get("/dashboard/{session_id}", response_model=DashboardResponse)
async def dashboard(
    session_id: str,
    granularity: Optional[str] = Query(None),
    filters: Optional[str] = Query(None),  # JSON-encoded dict e.g. '{"Customer Type":"Software"}'
    top_n: int = Query(10, ge=1, le=200),
):
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

    parsed_filters = None
    if filters:
        try:
            parsed_filters = json.loads(filters)
        except json.JSONDecodeError:
            raise HTTPException(400, "Invalid filters parameter — must be a JSON object.")

    loop = asyncio.get_event_loop()
    try:
        fn = partial(compute_dashboard, session_dir, filepath, granularity, parsed_filters, top_n)
        result = await loop.run_in_executor(None, fn)
    except Exception as e:
        raise HTTPException(500, f"Dashboard computation failed: {e}")

    return result
