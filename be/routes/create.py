from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from motor.core import AgnosticDatabase
from be.db import get_db

router = APIRouter()

class CreateCollectionRequest(BaseModel):
    collection: str

@router.post("/create_collection")
async def create_collection(
    request: CreateCollectionRequest,
    db: AgnosticDatabase = Depends(get_db)
):
    try:
        # Explicitly create the collection
        await db.create_collection(request.collection)
        
        return {"message": f"Collection '{request.collection}' created successfully!"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
