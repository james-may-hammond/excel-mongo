"""
Module: main.py
Description: Entry point for the FastAPI backend application. Registers all REST and WebSocket routes.
Dependencies: fastapi, motor, uvicorn
"""
import os
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from starletter.middleware.base import BaseHTTPMiddleware

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
from be.routes.auth import router as auth_router

ROOT_DIR = Path(__file__).resolve().parent.parent

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        response = await call_next(request)

        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["X-XSS-Protection"] = "1; mode=block"

        if os.getenv("ENVIRONMENT") == "production":
            response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
            
        return response

app = FastAPI(
    title="excel-mongo connector",
    version="1.1.0"
)
frontend_urls_env = os.getenv("FRONTEND_URLS", "https://localhost:3000")
allowed_origins = [url.strip() for url in frontend_urls_env.split(",") if url.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
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
app.include_router(auth_router)


app.include_router(ws_router)
app.add_middleware(SecurityHeadersMiddleware)

@app.get("/")
async def root():
    return {
        "service": "excel-mongo connector",
        "status": "running"
    }

