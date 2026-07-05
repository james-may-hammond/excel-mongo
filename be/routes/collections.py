"""
Module: collections.py
Description: API route to list all collections in the connected MongoDB database.
Dependencies: fastapi, motor
"""
from fastapi import APIRouter, HTTPException, Depends
from be.db import get_db
router = APIRouter()

@router.get("/collections")
async def collections(db = Depends(get_db)):
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