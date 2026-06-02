from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from be.db import mongo_client, db

router = APIRouter()
class FetchRequest(BaseModel):
    collection: str
    filters: dict = {}


@router.post("/fetch")
async def fetch(request: FetchRequest):
    try:
        collection = db[request.collection]

        records = await collection.find(request.filters).to_list()
        for record in records: record["_id"] = str(record["_id"])

        return {
            "status": "success",
            "count": len(records),
            "data": records
        }
    except Exception as e:
        raise HTTPException (
            status_code=500,
            detail=str(e)
        )