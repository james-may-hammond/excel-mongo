"""
Module: db.py
Description: Handles MongoDB connection initialization, caching, and dependency injection.
Dependencies: fastapi, motor
"""
from fastapi import Header, HTTPException
from motor.core import AgnosticDatabase
from motor.motor_asyncio import AsyncIOMotorClient

client_cache = {}

async def get_db(x_mongo_uri: str = Header(None), x_mongo_db: str = Header(None)) -> AgnosticDatabase:
    if not x_mongo_uri or not x_mongo_db:
        raise HTTPException(
            status_code=401,
            detail="Missing X-Mongo-URI or X-Mongo-DB headers. Please login in the Add-in."
        )
    return await get_ws_db(x_mongo_uri, x_mongo_db)

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
