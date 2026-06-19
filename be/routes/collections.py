from fastapi import APIRouter, HTTPException, Depends
from be.db import get_db
from motor.core import AgnosticDatabase
router = APIRouter()

@router.get("/collections")
async def collections(db: AgnosticDatabase = Depends(get_db)):
    try:
        collections = await db.list_collection_names()
        return {
            "collections": collections
        }
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )