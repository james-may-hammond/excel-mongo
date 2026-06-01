from fastapi import APIRouter
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
        return {
            "status": "unhealthy",
            "error": str(e)
        }
    
    