"""
Module: health.py
Description: Health check API route to verify the service is running.
Dependencies: fastapi
"""
from fastapi import APIRouter

router = APIRouter()

@router.get("/health")
async def health():
    return {
        "status": "ok",
        "message": "Finnoto Proxy is running"
    }