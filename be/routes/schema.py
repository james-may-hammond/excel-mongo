"""
Module: schema.py
Description: API route to infer and retrieve the schema of a MongoDB collection.
Dependencies: fastapi, motor
"""
from fastapi import APIRouter, HTTPException, Query, Depends
from be.db import get_db
from motor.core import AgnosticDatabase

router = APIRouter()

@router.get("/schema")
async def get_schema(collection: str = Query(...), db: AgnosticDatabase = Depends(get_db)):
    try:
        col = db[collection]
        total_count = await col.estimated_document_count()
        docs = await col.find({}).limit(5).to_list(5)
        if not docs:
            return {
                "fields": [], 
                "total_count": total_count,
                "message": "Collection is empty — no schema could be inferred."
            }

        field_set = ["_id"]
        for doc in docs:
            for key in doc.keys():
                if key not in ["_id", "__v"] and key not in field_set:
                    field_set.append(key)

        return {
            "fields": field_set,
            "sampled": len(docs),
            "total_count": total_count
        }
    
    except Exception as e:
        raise HTTPException(
            status_code=500, 
            detail=str(e)
        )
