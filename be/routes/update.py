from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from be.db import db
from bson import ObjectId

router = APIRouter()

class UpdateRequest(BaseModel):
    collection: str
    data: dict
    id: str

@router.post("/update")
async def update(request: UpdateRequest):
    try:
        collection = db[request.collection]
        try:
            obj_id = ObjectId(request.id)
        except Exception:
            obj_id = request.id
        data = request.data.copy()
        if "_id" in data:
            data.pop("_id")

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