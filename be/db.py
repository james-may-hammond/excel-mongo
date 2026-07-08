"""
Module: db.py
Description: Handles Finnoto REST API proxying wrapped in a MongoDB-like client interface.
             All write operations (insert, update, delete) are stubs that return 405
             since Finnoto data is read-only via this proxy.
"""
import httpx
import json
from typing import Any, Dict, List
from fastapi import Header, HTTPException

# ---------------------------------------------------------------------------
# Endpoint registry
# Each entry maps a "collection" key to the Finnoto API URL + required payload.
# The "label" is shown in the Excel Add-in collection dropdown.
# ---------------------------------------------------------------------------
FINNOTO_ENDPOINTS: Dict[str, Dict[str, Any]] = {
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

# Shared headers sent with every Finnoto API call
FINNOTO_COMMON_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "Origin": "https://devfn.vercel.app",
    "Referer": "https://devfn.vercel.app/",
    "u-device-version": "1.0.0",
    "u-platform-id": "4",
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36"
    )
}


def flatten_json(y: Any) -> Dict[str, Any]:
    """Recursively flatten a nested JSON object into dot-notation keys."""
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
    """
    Extract the record list from various Finnoto response shapes:
      - { data: [...] }
      - { data: { items: [...] } }
      - { data: { list: [...] } }
      - { data: { data: [...] } }
      - { items: [...] }
      - { list: [...] }
      - Any top-level key whose value is a list
    """
    data = response_json.get("data")
    if data is not None:
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            for key in ("items", "list", "data", "records", "results"):
                if key in data and isinstance(data[key], list):
                    return data[key]

    for key in ("items", "list", "records", "results"):
        if key in response_json and isinstance(response_json[key], list):
            return response_json[key]

    # Last-resort: first list value found at the top level
    for val in response_json.values():
        if isinstance(val, list):
            return val

    return []


class FinnotoCollection:
    def __init__(self, name: str, token: str):
        self.name = name
        self.token = token

    def _build_headers(self) -> dict:
        return {**FINNOTO_COMMON_HEADERS, "Authorization": f"Bearer {self.token}"}

    async def _fetch_records(self, filter: dict, limit: int) -> list:
        if self.name not in FINNOTO_ENDPOINTS:
            raise HTTPException(
                status_code=400,
                detail=f"Unsupported collection: '{self.name}'. "
                       f"Available: {list(FINNOTO_ENDPOINTS.keys())}"
            )

        config = FINNOTO_ENDPOINTS[self.name]
        url = config["url"]
        payload = config["payload"].copy()

        # Inject dynamic query filters passed from the client
        for k, v in filter.items():
            payload[k] = v

        if limit > 0:
            payload["limit"] = limit

        async with httpx.AsyncClient(timeout=30.0) as client:
            try:
                res = await client.post(url, json=payload, headers=self._build_headers())
            except Exception as e:
                raise HTTPException(
                    status_code=503,
                    detail=f"Failed to connect to Finnoto API: {str(e)}"
                )

            if res.status_code in (401, 403):
                raise HTTPException(
                    status_code=401,
                    detail="Finnoto token is invalid or expired. Please reconnect."
                )
            if res.status_code != 200:
                raise HTTPException(
                    status_code=502,
                    detail=f"Finnoto API error (HTTP {res.status_code}): {res.text[:400]}"
                )

            try:
                resp_data = res.json()
            except Exception:
                raise HTTPException(status_code=502, detail="Finnoto returned non-JSON response.")

        raw_records = extract_records(resp_data)
        flattened = []
        for index, item in enumerate(raw_records):
            flat = flatten_json(item)
            if "_id" not in flat:
                flat["_id"] = str(flat.get("id", f"row_{index + 1}"))
            flattened.append(flat)

        return flattened

    # ------------------------------------------------------------------
    # MockCursor — wraps the async fetch in a cursor-like interface
    # ------------------------------------------------------------------
    class MockCursor:
        def __init__(self, coro, limit_val=0):
            self.coro = coro
            self.limit_val = limit_val
            self.records = None
            self.index = 0

        def limit(self, val):
            self.limit_val = val
            return self

        def skip(self, count):
            # Not meaningful for Finnoto; kept for API compatibility
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
            raise StopAsyncIteration

    # ------------------------------------------------------------------
    # Read interface
    # ------------------------------------------------------------------
    def find(self, filter: dict = None, projection: dict = None, limit: int = 0):
        """
        projection arg accepted for API compatibility but ignored —
        Finnoto returns fixed response shapes.
        """
        coro = self._fetch_records(filter or {}, limit)
        return self.MockCursor(coro, limit)

    async def estimated_document_count(self):
        # Finnoto doesn't expose a count endpoint; return a placeholder
        return 1000

    # ------------------------------------------------------------------
    # Write stubs — Finnoto is read-only
    # ------------------------------------------------------------------
    def _read_only(self):
        raise HTTPException(
            status_code=405,
            detail="Finnoto data is read-only. Write operations are not supported."
        )

    async def insert_one(self, document: dict):
        self._read_only()

    async def insert_many(self, documents: list):
        self._read_only()

    async def update_many(self, filter: dict, update: dict, upsert: bool = False):
        self._read_only()

    async def replace_one(self, filter: dict, replacement: dict, upsert: bool = False):
        self._read_only()

    async def find_one(self, filter: dict):
        self._read_only()

    async def bulk_write(self, operations, ordered=False):
        self._read_only()

    async def delete_many(self, filter: dict):
        self._read_only()


class FinnotoDatabase:
    def __init__(self, token: str):
        self.token = token

    def __getitem__(self, name: str) -> FinnotoCollection:
        return FinnotoCollection(name, self.token)

    async def list_collection_names(self) -> List[dict]:
        """Return endpoint metadata (key + label) for the dropdown."""
        return [
            {"key": key, "label": cfg["label"]}
            for key, cfg in FINNOTO_ENDPOINTS.items()
        ]

    async def create_collection(self, name: str):
        raise HTTPException(
            status_code=405,
            detail="Finnoto endpoints are predefined. Custom collections cannot be created."
        )

    async def command(self, cmd: dict):
        return {"ok": 1.0}


async def get_db(authorization: str = Header(None)):
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=401,
            detail="Missing or invalid Authorization header. Please reconnect in the Add-in."
        )
    token = authorization[len("Bearer "):].strip()
    if not token:
        raise HTTPException(status_code=401, detail="Empty Bearer token.")
    return FinnotoDatabase(token)


async def get_ws_db(token: str):
    if not token:
        raise HTTPException(status_code=401, detail="Missing token for WebSocket connection.")
    return FinnotoDatabase(token)
