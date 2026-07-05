"""
Module: db.py
Description: Handles Finnoto REST API proxying wrapped in a MongoDB-like client interface.
"""
import httpx
import json
from typing import Any, Dict, List
from fastapi import Header, HTTPException

FINNOTO_ENDPOINTS = {
    "ap.pending.approval.report.list": {
        "label": "Pending Approval Reports",
        "url": "https://eapi.finnoto.dev/api/b/report/ap.pending.approval.report.list/list",
        "payload": {
            "slug": "ap.pending.approval.report.list",
            "source_type": "9ff68c2947fabaef5b67608904595189",
            "listing_slug": "ap_pending_approval_reports_modal",
            "document_type_identifier": "fetch_report",
            "ignore_dto_all": True
        }
    },
    "expense": {
        "label": "Employee Expenses",
        "url": "https://eapi.finnoto.dev/api/b/expense/search",
        "payload": {
            "is_draft": False,
            "listing_slug": "ef_expenses",
            "document_type_identifier": "employee_expenses",
            "ignore_dto_all": True
        }
    }
}

def flatten_json(y: Any) -> Dict[str, Any]:
    out = {}
    def flatten(x: Any, name: str = ''):
        if isinstance(x, dict):
            for a in x:
                flatten(x[a], name + a + '.')
        elif isinstance(x, list):
            for i, a in enumerate(x):
                flatten(a, name + str(i) + '.')
        else:
            out[name[:-1]] = x
    flatten(y)
    return out

def extract_records(response_json: dict) -> list:
    if "data" in response_json:
        if isinstance(response_json["data"], list):
            return response_json["data"]
        elif isinstance(response_json["data"], dict) and "items" in response_json["data"]:
            return response_json["data"]["items"]
    
    if "items" in response_json:
        return response_json["items"]
        
    for val in response_json.values():
        if isinstance(val, list):
            return val
            
    return []

class FinnotoCollection:
    def __init__(self, name: str, token: str):
        self.name = name
        self.token = token

    async def _fetch_records(self, filter: dict, limit: int):
        if self.name not in FINNOTO_ENDPOINTS:
            raise HTTPException(status_code=400, detail=f"Unsupported endpoint/collection: '{self.name}'")
            
        config = FINNOTO_ENDPOINTS[self.name]
        url = config["url"]
        payload = config["payload"].copy()
        
        # Inject dynamic query filters
        for k, v in filter.items():
            payload[k] = v
            
        if limit > 0:
            payload["limit"] = limit

        headers = {
            "Accept": "application/json, text/plain, */*",
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
            "Origin": "https://devfn.vercel.app",
            "Referer": "https://devfn.vercel.app/",
            "u-device-version": "1.0.0",
            "u-platform-id": "4",
            "User-Agent": "Mozilla/5.0"
        }
        
        async with httpx.AsyncClient(timeout=30.0) as client:
            try:
                res = await client.post(url, json=payload, headers=headers)
                if res.status_code != 200:
                    raise Exception(f"Finnoto API error: {res.text}")
                resp_data = res.json()
            except Exception as e:
                raise Exception(f"Failed to connect to Finnoto API: {str(e)}")

        raw_records = extract_records(resp_data)
        flattened_records = []
        for index, item in enumerate(raw_records):
            flat = flatten_json(item)
            if "_id" not in flat:
                flat["_id"] = str(flat.get("id", f"row_{index + 1}"))
            flattened_records.append(flat)
            
        return flattened_records

    class MockCursor:
        def __init__(self, coro, limit):
            self.coro = coro
            self.limit_val = limit
            self.records = None
            self.index = 0
            
        def limit(self, val):
            self.limit_val = val
            return self
            
        def skip(self, count):
            return self

        def batch_size(self, size):
            return self
            
        async def to_list(self, length):
            if self.records is None:
                self.records = await self.coro
            return self.records[:length]
            
        def __aiter__(self):
            return self
            
        async def __anext__(self):
            if self.records is None:
                self.records = await self.coro
            if self.index < len(self.records):
                item = self.records[self.index]
                self.index += 1
                return item
            else:
                raise StopAsyncIteration
                
    def find(self, filter: dict = None, limit: int = 0):
        coro = self._fetch_records(filter or {}, limit)
        return self.MockCursor(coro, limit)
        
    async def estimated_document_count(self):
        # Stub for schema inference
        return 1000

    async def insert_many(self, documents: list):
        class InsertResult:
            def __init__(self, count):
                self.inserted_ids = [str(i) for i in range(count)]
        return InsertResult(len(documents))
        
    async def update_many(self, filter: dict, update: dict, upsert: bool = False):
        class UpdateResult:
            def __init__(self):
                self.modified_count = 1
                self.upserted_id = None
        return UpdateResult()
        
    async def bulk_write(self, operations, ordered=False):
        class BulkWriteResult:
            def __init__(self, count):
                self.matched_count = count
        return BulkWriteResult(len(operations))
        
    async def delete_many(self, filter: dict):
        class DeleteResult:
            def __init__(self):
                self.deleted_count = 1
        return DeleteResult()

class FinnotoDatabase:
    def __init__(self, token: str):
        self.token = token
        
    def __getitem__(self, name: str) -> FinnotoCollection:
        return FinnotoCollection(name, self.token)
        
    async def list_collection_names(self):
        return list(FINNOTO_ENDPOINTS.keys())
        
    async def create_collection(self, name: str):
        # Stub for creating new endpoints
        pass
        
    async def command(self, cmd: dict):
        return {"ok": 1.0}

async def get_db(authorization: str = Header(None)):
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=401,
            detail="Missing or invalid Authorization header. Please login in the Add-in."
        )
    token = authorization.replace("Bearer ", "").strip()
    return FinnotoDatabase(token)

async def get_ws_db(token: str):
    return FinnotoDatabase(token)
