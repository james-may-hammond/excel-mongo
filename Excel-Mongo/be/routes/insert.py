"""
Module: insert.py
Description: Insert route — returns 405 since Finnoto data is read-only.
"""
from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel

router = APIRouter()

class InsertionRequest(BaseModel):
    collection: str
    data: dict

@router.post("/insert", status_code=status.HTTP_201_CREATED)
async def insert(request: InsertionRequest):
    raise HTTPException(
        status_code=405,
        detail="Finnoto data is read-only. Insert operations are not supported."
    )
