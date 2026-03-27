"""
FastAPI entry point for the Data Cutter web application.
"""

import asyncio
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from routers.wizard import router as wizard_router
from routers.dashboard import router as dashboard_router
from temp_store import cleanup_old_sessions


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Startup / shutdown lifecycle."""
    # Start background cleanup task
    task = asyncio.create_task(_periodic_cleanup())
    yield
    task.cancel()


async def _periodic_cleanup():
    """Clean up stale sessions every 15 minutes."""
    while True:
        await asyncio.sleep(900)
        cleanup_old_sessions(max_age_seconds=3600)


app = FastAPI(
    title="Data Cutter",
    description="Upload raw ARR data, configure analysis, download formatted Excel.",
    lifespan=lifespan,
)

# CORS — allow Vite dev server
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(wizard_router)
app.include_router(dashboard_router)
