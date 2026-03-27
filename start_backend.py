"""
Backend launcher for preview_start compatibility.
Removes restricted Desktop paths from sys.path before uvicorn imports them,
then adds the correct paths explicitly.
"""
import sys

# Strip any Desktop or cwd paths that the sandbox can't stat
sys.path = [p for p in sys.path if "Desktop" not in p and p != ""]

# Add user site-packages (where pip --user installs packages)
sys.path.insert(0, "/Users/sammichaels/Library/Python/3.9/lib/python/site-packages")

# Add the backend source directory so 'main' is importable
sys.path.insert(0, "/Users/sammichaels/Desktop/Coding/Data Cutter/backend")

import uvicorn  # noqa: E402 — must come after sys.path setup

uvicorn.run(
    "main:app",
    host="127.0.0.1",
    port=8000,
    loop="asyncio",
    http="h11",
)
