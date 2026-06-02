from fastapi import APIRouter, HTTPException
from be.db import mongo_client

router = APIRouter()

@router.get("/health")
async def health():
    try:
        await mongo_client.admin.command("ping")
        return {
            "status": "ok"
        }
    except Exception as e:
        raise HTTPException (
            status_code=500,
            detail=str(e)
        )
    
    