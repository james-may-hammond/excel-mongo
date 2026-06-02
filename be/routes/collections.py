from fastapi import APIRouter, HTTPException
from be.db import mongo_client,db
router = APIRouter()

@router.get("/collections")
async def collections():
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