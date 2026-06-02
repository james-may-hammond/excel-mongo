from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from be.db import mongo_client, db

router = APIRouter()
class InsertionRequest(BaseModel):
    collection: str
    data: dict


@router.post("/insert", status_code=status.HTTP_201_CREATED)
async def insert(request: InsertionRequest):
    try:
        collection = db[request.collection]
        
        await collection.insert_one (
            request.data
        )
        return {
            "status":"sucess"
        }
    
    except Exception as e:
        raise HTTPException (
            status_code=501,
            detail=str(e)
        )
