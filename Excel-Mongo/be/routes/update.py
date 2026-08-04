"""
Module: update.py
Description: Update route — returns 405 since Finnoto data is read-only.
"""
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

router = APIRouter()

class UpdateRequest(BaseModel):
    collection: str
    data: dict
    id: str

@router.post("/update")
async def update(request: UpdateRequest):
    raise HTTPException(
        status_code=405,
        detail="Finnoto data is read-only. Update operations are not supported."
    )