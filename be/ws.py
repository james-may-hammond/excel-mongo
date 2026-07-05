"""
Module: ws.py
Description: WebSocket connection manager and event router for real-time Excel-Mongo synchronization.
Dependencies: fastapi, motor, bson, json
"""
import asyncio
import base64
import json
from datetime import datetime

from bson import ObjectId
from fastapi import APIRouter, WebSocket, WebSocketDisconnect

from be.db import get_ws_db

# Existing Routers
from be.routes.bulk import (
    BulkInsertRequest,
    BulkUpdateRequest,
    bulk_insert,
    bulk_update,
)
from be.routes.collections import collections
from be.routes.create import CreateCollectionRequest, create_collection
from be.routes.delete import BulkDeleteRequest, bulk_delete
from be.routes.fetch import FetchRequest as FetchReq
from be.routes.fetch import fetch
from be.routes.insert import InsertionRequest, insert
from be.routes.schema import get_schema
from be.routes.update import UpdateRequest as UpdateReq
from be.routes.update import update


def mongo_serializer(obj):
    if isinstance(obj, ObjectId):
        return str(obj)
    if isinstance(obj, datetime):
        return obj.isoformat()
    if isinstance(obj, bytes):
        return base64.b64encode(obj).decode('utf-8')
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")

router = APIRouter()

class ConnectionManager:
    def __init__(self):
        self.active_connections : list[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def send_response(self, websocket: WebSocket, request_id: str, status: str, data: dict = None, error: str = None):
         response = {"requestID": request_id, "status": status}
         if data is not None: response["data"] = data
         if error is not None: response["error"] = error
         await websocket.send_text(json.dumps(response))

manager = ConnectionManager()

@router.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket, token: str):
    await manager.connect(websocket)
    try:
        db = await get_ws_db(token)
    except Exception as e:
        await manager.send_response(websocket, "auth", "error", error="Invalid token or connection failed")
        await websocket.close()
        return
    try:
        while True:
            text = await websocket.receive_text()
            try: msg = json.loads(text)
            except json.JSONDecodeError: continue
            request_id = msg.get("requestId")
            action = msg.get("action")
            payload = msg.get("payload", {})
            if not request_id or not action: continue
            try:
                if action == "collections":
                    res = await collections(db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                elif action == "schema":
                    res = await get_schema(collection=payload.get("collection"), db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                elif action == "fetch":
                    res = await fetch(request=FetchReq(**payload), db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                elif action == "stream_fetch":
                    col = db[payload.get("collection")]
                    cursor = col.find(payload.get("filters", {}))
                    if payload.get("limit", 0) > 0: cursor = cursor.limit(payload.get("limit", 0))

                    count = 0
                    try:
                        async for doc in cursor.batch_size(1000):
                            doc["_id"] = str(doc["_id"])
                            doc.pop("__v", None)
                            await websocket.send_text(json.dumps({
                                "requestId": request_id, "status": "chunk", "data": doc
                            }, default=mongo_serializer))

                            count += 1
                            if count % 2000 == 0: await asyncio.sleep(0.5)
                        await manager.send_response(websocket, request_id, "success", data={"message": "Stream complete"})
                    except Exception as e:
                        await manager.send_response(websocket, request_id, "error", error=str(e))
                elif action == "insert":
                    res = await insert(request=InsertionRequest(**payload), db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                elif action == "update":
                    res = await update(request=UpdateReq(**payload), db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                elif action == "bulk_insert":
                    res = await bulk_insert(request=BulkInsertRequest(**payload), db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                elif action == "bulk_update":
                    res = await bulk_update(request=BulkUpdateRequest(**payload), db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                elif action == "bulk_delete":
                    res = await bulk_delete(requests=BulkDeleteRequest(**payload), db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                elif action == "create_collection":
                    res = await create_collection(request=CreateCollectionRequest(**payload), db=db)
                    await manager.send_response(websocket, request_id, "success", data=res)

                else:
                    await manager.send_response(websocket, request_id, "error", error=f"Unknown action: {action}")

            except Exception as e:
                # Catch any HTTPException raised by the routes
                await manager.send_response(websocket, request_id, "error", error=getattr(e, "detail", str(e)))
    except WebSocketDisconnect:
        manager.disconnect(websocket)
