from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from be.db import mongo_client

router = APIRouter()
class InsertionRequest(BaseModel):
    collection: str
    data: dict


@router.post("/insert")
async def insert(request: InsertionRequest):
    try:
        db = mongo_client["excel_mongo"]
        collection = db[request.collection]
        
        await collection.insert_one (
            request.data
        )
        return {
            "status":"sucess"
        }
    
    except Exception as e:
        raise HTTPException (
            status_code=501
            detail=str(e)
        )
