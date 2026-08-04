"""
Module: db.py
Description: Finnoto API client.
             FinnotoDatabase is the main interface — it proxies requests to the
             Finnoto REST API using a per-request Bearer token.
             All write operations return 405 (read-only proxy).
"""
import httpx
from typing import Any, Dict, List
from fastapi import Header, HTTPException

# Default Finnoto API base (can be overridden at login time)
# Preset environments:
#   eapi     → https://eapi.finnoto.dev
#   abdebug  → https://abdebug.finnoto.dev
#   arcapi   → https://arcapi.finnoto.dev
import os
DEFAULT_FINNOTO_BASE = os.getenv("FINNOTO_BASE", "https://eapi.finnoto.dev")

# Shared headers sent with every Finnoto call
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


# ---------------------------------------------------------------------------
# JSON helpers
# ---------------------------------------------------------------------------
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
      - { data: { items/list/data/records/results: [...] } }
      - { items/list/records/results: [...] }
      - first list value at top level
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

    for val in response_json.values():
        if isinstance(val, list):
            return val

    return []


# ---------------------------------------------------------------------------
# FinnotoCollection — cursor-style interface over a single Finnoto endpoint
# ---------------------------------------------------------------------------

class FinnotoCollection:
    """
    Wraps a dynamic Finnoto endpoint (resolved at query time from the report API)
    with a cursor-like interface for compatibility with the existing fetch routes.
    """
    def __init__(self, name: str, token: str, base_url: str, endpoint_config: dict):
        self.name = name
        self.token = token
        self.base_url = base_url
        self._config = endpoint_config   # { url, payload }

    def _build_headers(self) -> dict:
        return {**FINNOTO_COMMON_HEADERS, "Authorization": f"Bearer {self.token}"}

    async def _fetch_records(self, filter: dict, limit: int) -> list:
        url = self._config["url"]
        payload = self._config.get("payload", {}).copy()

        for k, v in filter.items():
            payload[k] = v
        if limit > 0:
            payload["limit"] = limit

        async with httpx.AsyncClient(timeout=30.0) as client:
            try:
                res = await client.post(url, json=payload, headers=self._build_headers())
            except Exception as e:
                raise HTTPException(status_code=503, detail=f"Failed to reach Finnoto: {e}")

            if res.status_code in (401, 403):
                raise HTTPException(status_code=401, detail="Finnoto token is invalid or expired. Please reconnect.")
            if res.status_code != 200:
                raise HTTPException(status_code=502, detail=f"Finnoto API error (HTTP {res.status_code}): {res.text[:400]}")

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
    # MockCursor
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

    def find(self, filter: dict = None, projection: dict = None, limit: int = 0):
        coro = self._fetch_records(filter or {}, limit)
        return self.MockCursor(coro, limit)

    async def estimated_document_count(self):
        return 1000

    # ------------------------------------------------------------------
    # Write stubs — Finnoto is read-only
    # ------------------------------------------------------------------
    def _read_only(self):
        raise HTTPException(status_code=405, detail="Finnoto data is read-only. Write operations are not supported.")

    async def insert_one(self, document: dict):        self._read_only()
    async def insert_many(self, documents: list):      self._read_only()
    async def update_many(self, *a, **kw):             self._read_only()
    async def replace_one(self, *a, **kw):             self._read_only()
    async def find_one(self, filter: dict):            self._read_only()
    async def bulk_write(self, operations, **kw):      self._read_only()
    async def delete_many(self, filter: dict):         self._read_only()


# ---------------------------------------------------------------------------
# FinnotoDatabase — top-level interface
# ---------------------------------------------------------------------------

class FinnotoDatabase:
    """
    Represents an authenticated Finnoto session.
    Collection lookups are dynamic — the collection name is resolved against the
    live report list if not a well-known built-in key.
    """
    def __init__(self, token: str, base_url: str = DEFAULT_FINNOTO_BASE):
        self.token = token
        self.base_url = base_url.rstrip("/")

    def __getitem__(self, name: str) -> FinnotoCollection:
        """
        Returns a FinnotoCollection for the given endpoint name.
        The endpoint config is resolved dynamically; callers pass the
        full URL + payload via the _endpoint_config override, or the
        collection is treated as a generic search endpoint.
        """
        # Generic fallback — unknown collections just get an empty cursor
        config = {"url": f"{self.base_url}/api/b/{name}/search", "payload": {"ignore_dto_all": True}}
        return FinnotoCollection(name, self.token, self.base_url, config)

    async def list_collection_names(self) -> List[dict]:
        """
        Dynamically fetch available business reports from Finnoto.
        Returns [{ key, label }] for the dropdown.
        """
        import httpx
        payload = {"page": 1}
        headers = {**FINNOTO_COMMON_HEADERS, "Authorization": f"Bearer {self.token}"}
        try:
            async with httpx.AsyncClient(timeout=15.0) as client:
                res = await client.post(
                    f"{self.base_url}/api/b/report/search",
                    json=payload,
                    headers=headers
                )
            if res.status_code in (401, 403):
                raise HTTPException(status_code=401, detail="Finnoto token expired. Please reconnect.")
            if res.status_code >= 400:
                return []
            raw = res.json()
            reports = extract_records(raw)
            return [
                {"key": str(r.get("id", i)), "label": r.get("name") or r.get("title") or f"Report {i+1}"}
                for i, r in enumerate(reports)
            ]
        except HTTPException:
            raise
        except Exception:
            return []

    async def create_collection(self, name: str):
        raise HTTPException(status_code=405, detail="Finnoto endpoints are predefined.")

    async def command(self, cmd: dict):
        return {"ok": 1.0}


# ---------------------------------------------------------------------------
# FastAPI dependency helpers
# ---------------------------------------------------------------------------

async def get_db(
    authorization: str = Header(None),
    x_finnoto_base: str = Header(None, alias="X-Finnoto-Base")
) -> FinnotoDatabase:
    """
    Extracts the Bearer token from the Authorization header.
    Optionally accepts X-Finnoto-Base header to override the API base URL
    (useful for pointing at debug/staging instances).
    """
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=401,
            detail="Missing or invalid Authorization header. Please reconnect in the Add-in."
        )
    token = authorization[len("Bearer "):].strip()
    if not token:
        raise HTTPException(status_code=401, detail="Empty Bearer token.")

    return FinnotoDatabase(token, x_finnoto_base or DEFAULT_FINNOTO_BASE)


async def get_ws_db(token: str, base_url: str = DEFAULT_FINNOTO_BASE) -> FinnotoDatabase:
    if not token:
        raise HTTPException(status_code=401, detail="Missing token for WebSocket connection.")
    return FinnotoDatabase(token, base_url or DEFAULT_FINNOTO_BASE)
