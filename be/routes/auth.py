"""
Module: auth.py
Description: Authentication route that exchanges MongoDB credentials for a securely encrypted access token.
Dependencies: fastapi, motor, pydantic, cryptography, json, os
"""
import os
import json
from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from motor.motor_asyncio import AsyncIOMotorClient

from cryptography.fernet import Fernet

router = APIRouter()
import base64
SECRET_KEY = os.getenv("SECRET_KEY", "your-fallback-dev-secret-key-that-is-32-bytes")
FERNET_KEY = SECRET_KEY.ljust(32, '0')[:32].encode('utf-8')
FERNET_KEY = base64.urlsafe_b64encode(FERNET_KEY)
cipher_suite = Fernet(FERNET_KEY)

class LoginRequest(BaseModel):
    uri: str
    db_name: str

@router.post("/auth/connect")
async def connect(request: LoginRequest):
    try:
        client = AsyncIOMotorClient(request.uri, serverSelectionTimeoutMS=3000)
        
        # Validate database name casing
        try:
            db_names = await client.list_database_names()
            if request.db_name not in db_names:
                for existing_db in db_names:
                    if existing_db.lower() == request.db_name.lower():
                        request.db_name = existing_db
                        break
        except Exception:
            # User might not have listDatabases privilege, fallback to provided name
            pass

        await client[request.db_name].list_collection_names()
        payload = json.dumps({"uri": request.uri, "db_name": request.db_name})
        encrypted_token = cipher_suite.encrypt(payload.encode('utf-8')).decode('utf-8')

        return {
            "status": "success", 
            "token": encrypted_token
        }
    except Exception as e:
        raise HTTPException(
            status_code=401,
            detail=str(e)
        )
    
    
    