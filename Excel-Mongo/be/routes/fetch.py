"""
Module: fetch.py
Description: API route to fetch documents from a collection with pagination.
Dependencies: fastapi, motor, pydantic
"""
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from be.db import get_db

router = APIRouter()
class FetchRequest(BaseModel):
    collection: str
    filters: dict = {}


@router.post("/fetch")
async def fetch(request: FetchRequest, db = Depends(get_db)):
    try:
        collection = db[request.collection]

        # Cap the maximum returned records to 1000 to keep Excel snappy
        records = await collection.find(request.filters).limit(1000).to_list(length=1000)
        for record in records: 
            record["_id"] = str(record["_id"])
            record.pop("__v", None)

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