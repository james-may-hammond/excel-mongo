"""
Module: collections.py
Description: API route to list all available Finnoto endpoint collections.
             Returns both the key (used internally) and the label (shown in the UI).
"""
from fastapi import APIRouter, HTTPException, Depends
from be.db import get_db

router = APIRouter()

@router.get("/collections")
async def collections(db=Depends(get_db)):
    try:
        items = await db.list_collection_names()
        return {"collections": items}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))