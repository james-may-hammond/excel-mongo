"""
Module: health.py
Description: Health check API route to verify the service is running.
Dependencies: fastapi
"""
from fastapi import APIRouter, Header, HTTPException
from be.db import client_cache
from motor.motor_asyncio import AsyncIOMotorClient

router = APIRouter()

@router.get("/health")
async def health(x_mongo_uri: str = Header(None), x_mongo_db: str = Header(None)):
    if not x_mongo_uri or not x_mongo_db:
        # Generic health check for Railway/Cloud providers
        return {
            "status": "ok",
            "message": "Server is running (no DB headers provided)"
        }
    
    # Frontend ping check
    try:
        if x_mongo_uri not in client_cache:
            client_cache[x_mongo_uri] = AsyncIOMotorClient(x_mongo_uri, serverSelectionTimeoutMS=5000)
        client = client_cache[x_mongo_uri]
        db = client[x_mongo_db]
        await db.command("ping")
        return {
            "status": "ok"
        }
    except Exception as e:
        raise HTTPException (
            status_code=500,
            detail=str(e)
        )