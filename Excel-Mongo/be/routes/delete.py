"""
Module: delete.py
Description: API route for bulk deleting documents by their IDs.
Dependencies: fastapi, motor, pydantic, bson
"""
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from typing import List
from be.db import get_db
from bson import ObjectId

class BulkDeleteRequest(BaseModel):
    collection: str
    ids: List[str]

router = APIRouter()
@router.post("/bulk_delete")
async def bulk_delete(requests: BulkDeleteRequest, db = Depends(get_db)):
    try:
        collection = db[requests.collection]

        obj_ids = []
        for i in requests.ids:
            try: obj_ids.append(ObjectId(i))
            except: obj_ids.append(i)

        result = await collection.delete_many({"_id": {"$in": obj_ids}})

        return {
            "status": "success",
            "deleted_count": result.deleted_count
        }
    except Exception as e:
        raise HTTPException(
            status_code=501,
            detail=str(e)
        )