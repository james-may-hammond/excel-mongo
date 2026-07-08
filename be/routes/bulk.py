"""
Module: bulk.py
Description: Bulk operation routes.
             bulk_insert and bulk_update return 405 (Finnoto is read-only).
             bulk_delete is kept as a stub for potential future use.
"""
from fastapi import APIRouter, HTTPException, status, Depends
from pydantic import BaseModel
from typing import List
from be.db import get_db

router = APIRouter()


class BulkInsertRequest(BaseModel):
    collection: str
    data: List[dict]


class BulkUpdateRequest(BaseModel):
    collection: str
    data: List[dict]


class BulkDeleteRequest(BaseModel):
    collection: str
    ids: List[str]


@router.post("/bulk_insert", status_code=status.HTTP_201_CREATED)
async def bulk_insert(request: BulkInsertRequest):
    raise HTTPException(
        status_code=405,
        detail="Finnoto data is read-only. Bulk insert is not supported."
    )


@router.post("/bulk_update")
async def bulk_update(request: BulkUpdateRequest):
    raise HTTPException(
        status_code=405,
        detail="Finnoto data is read-only. Bulk update is not supported."
    )


@router.post("/bulk_delete")
async def bulk_delete(requests: BulkDeleteRequest, db=Depends(get_db)):
    """
    Stub — kept for potential future use.
    Currently Finnoto does not support delete operations via this proxy.
    """
    raise HTTPException(
        status_code=405,
        detail="Finnoto data is read-only. Bulk delete is not supported."
    )
