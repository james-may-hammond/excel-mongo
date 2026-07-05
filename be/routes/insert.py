"""
Module: insert.py
Description: API route to insert a single document into a collection.
Dependencies: fastapi, motor, pydantic
"""
from fastapi import APIRouter, HTTPException, status, Depends
from pydantic import BaseModel
from be.db import get_db

router = APIRouter()
class InsertionRequest(BaseModel):
    collection: str
    data: dict


@router.post("/insert", status_code=status.HTTP_201_CREATED)
async def insert(request: InsertionRequest, db = Depends(get_db)):
    try:
        collection = db[request.collection]
        doc = request.data.copy()
        doc["__v"] = 1
        
        await collection.insert_one (
            doc
        )
        return {
            "status":"sucess"
        }
    
    except Exception as e:
        raise HTTPException (
            status_code=501,
            detail=str(e)
        )
