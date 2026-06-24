"""
Module: main.py
Description: Entry point for the FastAPI backend application. Registers all REST and WebSocket routes.
Dependencies: fastapi, motor, uvicorn
"""
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from be.routes.bulk import router as bulk_router
from be.routes.collections import router as collections_router
from be.routes.create import router as create_router
from be.routes.delete import router as delete_router
from be.routes.fetch import router as fetch_router
from be.routes.health import router as health_router
from be.routes.insert import router as insertion_router
from be.routes.schema import router as schema_router
from be.routes.stream_fetch import router as stream_fetch_router
from be.routes.update import router as update_router
from be.ws import router as ws_router

ROOT_DIR = Path(__file__).resolve().parent.parent

app = FastAPI(
    title="excel-mongo connector",
    version="1.0.0"
)

# Enable CORS for Office Add-in environments (including Excel Online)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(health_router)
app.include_router(insertion_router)
app.include_router(collections_router)
app.include_router(fetch_router)
app.include_router(update_router)
app.include_router(schema_router)
app.include_router(stream_fetch_router)
app.include_router(bulk_router)
app.include_router(delete_router)
app.include_router(create_router)

app.include_router(ws_router)

@app.get("/")
async def root():
    return {
        "service": "excel-mongo connector",
        "status": "running"
    }
