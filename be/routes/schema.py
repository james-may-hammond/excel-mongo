from fastapi import APIRouter, HTTPException, Query
from be.db import db

router = APIRouter()

@router.get("/schema")
async def get_schema(collection: str = Query(...)):
    try:
        col = db[collection]
        docs = await col.find({}).limit(5).to_list(5)
        if not docs:
            return {
                "fields": [], 
                "message": "Collection is empty — no schema could be inferred."
            }

        field_set = ["_id"]
        for doc in docs:
            for key in doc.keys():
                if key != "_id" and key not in field_set:
                    field_set.append(key)

        return {
            "fields": field_set,
            "sampled": len(docs)
        }
    
    except Exception as e:
        raise HTTPException(
            status_code=500, 
            detail=str(e)
        )
