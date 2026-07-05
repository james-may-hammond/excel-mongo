"""
Module: update.py
Description: API route to update a single document in a collection with conflict detection.
Dependencies: fastapi, motor, pydantic, bson
"""
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from be.db import get_db
from bson import ObjectId

router = APIRouter()

class UpdateRequest(BaseModel):
    collection: str
    data: dict
    id: str

@router.post("/update")
async def update(request: UpdateRequest, db = Depends(get_db)):
    collection = db[request.collection]
    try:
        obj_id = ObjectId(request.id)
    except Exception:
        obj_id = request.id

    data = request.data.copy()
    if "_id" in data:
        data.pop("_id")

    # Update - Conflict Resolution Method
    client_v = data.get("__v", 0)
    try:
        client_v = int(client_v)
    except:
        client_v = 0
    try:
        existing = await collection.find_one({"id": obj_id})
        if existing:
            server_v = existing.get("__v", 0)
            if client_v < server_v:
                raise HTTPException(
                    status_code=409,
                    detail="conflict, server version is newer."
                )
        data["__v"] = client_v + 1

        await collection.replace_one(
            {"_id": obj_id},
            data,
            upsert=True
        )
        return {"status": "success"}
    except Exception as e:
        raise HTTPException (
            status_code=501,
            detail=str(e)
        )