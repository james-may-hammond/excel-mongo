"""
Module: bulk.py
Description: API routes for bulk inserting and bulk updating documents in MongoDB.
Dependencies: fastapi, motor, pydantic, pymongo, bson
"""
from fastapi import APIRouter, HTTPException, status, Depends
from pydantic import BaseModel
from be.db import get_db
from bson import ObjectId
from pymongo import UpdateOne
from typing import List

router = APIRouter()

class BulkInsertRequest(BaseModel):
    collection: str
    data: List[dict]

class BulkUpdateRequest(BaseModel):
    collection: str
    data: List[dict]

@router.post("/bulk_insert", status_code=status.HTTP_201_CREATED)
async def bulk_insert(request: BulkInsertRequest, db = Depends(get_db)):
    try:
        collection = db[request.collection]
        docs = []
        for d in request.data:
            doc = d.copy()
            doc.pop("_rowIndex", None)
            doc["__v"] = 1
            docs.append(doc)
            
        if docs:
            await collection.insert_many(docs)
            
        return {"status": "success", "inserted": len(docs)}
    except Exception as e:
        raise HTTPException(status_code=501, detail=str(e))

@router.post("/bulk_update")
async def bulk_update(request: BulkUpdateRequest, db = Depends(get_db)):
    try:
        collection = db[request.collection]
        
        operations = []
        doc_map = {}
        for d in request.data:
            data = d.copy()
            doc_id_str = data.pop("_id")
            try:
                obj_id = ObjectId(doc_id_str)
            except:
                obj_id = doc_id_str
                
            client_version = int(data.get("__v", 0))
            data["__v"] = client_version + 1
            
            filter_query = {"_id": obj_id}
            if client_version == 0:
                filter_query["$or"] = [{"__v": 0}, {"__v": {"$exists": False}}]
            else:
                filter_query["__v"] = client_version
                
            doc_map[doc_id_str] = {
                "rowIndex": data.pop("_rowIndex", None),
                "client_version": client_version
            }
                
            operations.append(
                UpdateOne(
                    filter_query,
                    {"$set": data}
                )
            )
            
        if not operations:
            return {"status": "success", "updated": 0, "conflicts": []}
            
        result = await collection.bulk_write(operations, ordered=False)
        
        conflicts = []
        if result.matched_count < len(operations):
            requested_ids = list(doc_map.keys())
            requested_obj_ids = []
            for i in requested_ids:
                try:
                    requested_obj_ids.append(ObjectId(i))
                except:
                    requested_obj_ids.append(i)
                
            db_docs = await collection.find({"_id": {"$in": requested_obj_ids}}, {"_id": 1, "__v": 1}).to_list(length=len(requested_ids))
            db_doc_map = {str(d["_id"]): d.get("__v", 0) for d in db_docs}
            
            for doc_id_str, info in doc_map.items():
                server_v = db_doc_map.get(doc_id_str)
                client_v = info["client_version"]
                if server_v is not None and server_v > client_v:
                    conflicts.append({
                        "_id": doc_id_str,
                        "_rowIndex": info["rowIndex"],
                        "server_v": server_v,
                        "client_v": client_v
                    })
        
        if conflicts:
            return {"status": "partial_success", "updated": result.modified_count, "conflicts": conflicts}
            
        return {"status": "success", "updated": result.modified_count, "conflicts": []}
        
    except Exception as e:
        raise HTTPException(status_code=501, detail=str(e))
