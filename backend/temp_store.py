"""
Temporary file / session management.

Each wizard session gets a UUID-based directory under the system temp folder.
Uploaded files and generated output are stored there.
A background task cleans up sessions older than 1 hour.
"""

import os
import shutil
import tempfile
import time
import uuid

TEMP_BASE = os.path.join(tempfile.gettempdir(), "data-cutter-sessions")


def _ensure_base():
    os.makedirs(TEMP_BASE, exist_ok=True)


def create_session() -> str:
    """Create a new session directory, return the session ID."""
    _ensure_base()
    session_id = uuid.uuid4().hex[:12]
    os.makedirs(os.path.join(TEMP_BASE, session_id), exist_ok=True)
    return session_id


def get_session_dir(session_id: str) -> str:
    """Return the absolute path to a session's directory."""
    return os.path.join(TEMP_BASE, session_id)


def get_upload_path(session_id: str) -> str:
    """Return the path where the uploaded file is stored."""
    return os.path.join(get_session_dir(session_id), "uploaded.xlsx")


def get_output_path(session_id: str) -> str:
    """Return the path where the generated output is stored."""
    return os.path.join(get_session_dir(session_id), "output.xlsx")


def session_exists(session_id: str) -> bool:
    return os.path.isdir(get_session_dir(session_id))


def cleanup_old_sessions(max_age_seconds: int = 3600):
    """Delete session dirs older than max_age_seconds."""
    _ensure_base()
    now = time.time()
    for name in os.listdir(TEMP_BASE):
        path = os.path.join(TEMP_BASE, name)
        if os.path.isdir(path):
            age = now - os.path.getmtime(path)
            if age > max_age_seconds:
                shutil.rmtree(path, ignore_errors=True)
