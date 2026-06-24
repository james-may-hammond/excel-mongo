"""
Module: stream_fetch.py
Description: API route for streaming documents from a collection to handle large datasets.
Dependencies: fastapi, motor, pydantic, json
"""
from fastapi import APIRouter, Depends
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from be.db import get_db
from motor.core import AgnosticDatabase
import json
import asyncio
from bson import ObjectId
from datetime import datetime

import base64

def mongo_serializer(obj):
    if isinstance(obj, ObjectId):
        return str(obj)
    if isinstance(obj, datetime):
        return obj.isoformat()
    if isinstance(obj, bytes):
        return base64.b64encode(obj).decode('utf-8')
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")

router = APIRouter()

class FetchRequest(BaseModel):
    collection: str
    filters: dict = {}
    limit: int = 0

@router.post("/stream_fetch")
async def stream_fetch(request: FetchRequest, db: AgnosticDatabase = Depends(get_db)):
    collection = db[request.collection]
    
    async def generate():
        count = 0
        try:
            cursor = collection.find(request.filters)
            if request.limit > 0:
                cursor = cursor.limit(request.limit)
                
            async for doc in cursor.batch_size(1000):
                doc["_id"] = str(doc["_id"])
                doc.pop("__v", None)
                yield json.dumps(doc, default=mongo_serializer) + "\n"
                
                count += 1
                if count % 2000 == 0:
                    await asyncio.sleep(0.5)
        except Exception as e:
            yield json.dumps({"_error": str(e)}) + "\n"

    return StreamingResponse(generate(), media_type="application/x-ndjson")
