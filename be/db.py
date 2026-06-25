"""
Module: db.py
Description: Handles MongoDB connection initialization, caching, and dependency injection.
Dependencies: fastapi, motor
"""

import json
from fastapi import Header, HTTPException
from motor.core import AgnosticDatabase
from motor.motor_asyncio import AsyncIOMotorClient

from be.routes.auth import cipher_suite


client_cache = {}

async def get_db(authorization: str = Header(None)) -> AgnosticDatabase:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=401,
            detail="Missing or invalid Authorization header. Please login in the Add-in."
        )
    
    token = authorization.replace("Bearer ", "")
    try:
        decrypted_bytes = cipher_suite.decrypt(token.encode('utf-8'))
        payload = json.loads(decrypted_bytes.decode('utf-8'))
        uri = payload["uri"]
        db_name = payload["db_name"]
        
    except Exception:
        raise HTTPException(
            status_code=401,
            detail="Invalid or expired token. Please log in again."
        )

    return await get_ws_db(uri, db_name)

async def get_ws_db(uri: str, db_name: str) -> AgnosticDatabase:
    if not uri or not db_name:
        raise ValueError("Missing MongoDB URI or DB Name")
    if uri not in client_cache:
        try:
            client_cache[uri] = AsyncIOMotorClient(uri, serverSelectionTimeoutMS=5000)
        except Exception as e:
            raise ValueError(f"Failed to initialize mongo client: {str(e)}")
    client = client_cache[uri]
    return client[db_name]
