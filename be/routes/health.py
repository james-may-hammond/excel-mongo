from fastapi import APIRouter, HTTPException, Depends
from be.db import get_db
from motor.core import AgnosticDatabase

router = APIRouter()

@router.get("/health")
async def health(db: AgnosticDatabase = Depends(get_db)):
    try:
        await db.command("ping")
        return {
            "status": "ok"
        }
    except Exception as e:
        raise HTTPException (
            status_code=500,
            detail=str(e)
        )
    
    