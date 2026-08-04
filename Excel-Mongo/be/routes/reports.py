import asyncio
import httpx
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from typing import Any, Dict, List, Optional
from be.db import get_db, FINNOTO_COMMON_HEADERS, DEFAULT_FINNOTO_BASE

router = APIRouter(prefix="/reports", tags=["reports"])


class ReportListRequest(BaseModel):
    search: Optional[str] = None
    limit: Optional[int] = None
    page: Optional[int] = 1


class DateRangeDto(BaseModel):
    min: str
    max: str


class DateFilterDto(BaseModel):
    range: Optional[DateRangeDto] = None
    today: Optional[bool] = None
    yesterday: Optional[bool] = None
    week: Optional[bool] = None
    last_week: Optional[bool] = None
    month: Optional[bool] = None
    last_month: Optional[bool] = None
    current_quarter: Optional[bool] = None
    previous_quarter: Optional[bool] = None
    year: Optional[bool] = None
    current_fy: Optional[bool] = None
    previous_fy: Optional[bool] = None
    financial_year: Optional[bool] = None
    fixed_month: Optional[str] = None


class ReportGenerateRequest(BaseModel):
    downloadable_report_id: int
    date: Optional[DateFilterDto] = None
    filters: Optional[Dict[str, Any]] = None
    comment: Optional[str] = None
    source_type: Optional[str] = None


class ReportSheetRequest(BaseModel):
    business_report_id: int
    sheet_ids: List[int]
    date: Optional[DateFilterDto] = None
    filters: Optional[Dict[str, Any]] = None


FINNOTO_BASE_PRESETS = {
    "eapi": "https://eapi.finnoto.dev",
    "abdebug": "https://abdebug.finnoto.dev",
    "arcapi": "https://arcapi.finnoto.dev",
}


def _build_headers(token: str) -> dict:
    return {**FINNOTO_COMMON_HEADERS, "Authorization": f"Bearer {token}"}


async def _post(url: str, payload: dict, token: str) -> dict:
    try:
        async with httpx.AsyncClient(timeout=60.0) as client:
            res = await client.post(url, json=payload, headers=_build_headers(token))
    except Exception as e:
        raise HTTPException(status_code=503, detail=f"Could not reach Finnoto API: {e}")

    if res.status_code in (401, 403):
        raise HTTPException(
            status_code=401,
            detail="Finnoto token is invalid or expired. Please reconnect."
        )
    if res.status_code == 404:
        raise HTTPException(status_code=404, detail="Not Found")
    if res.status_code >= 400:
        raise HTTPException(
            status_code=502,
            detail=f"Finnoto API error (HTTP {res.status_code}): {res.text[:400]}"
        )
    try:
        return res.json()
    except Exception:
        raise HTTPException(status_code=502, detail="Finnoto returned non-JSON response.")


async def _get(url: str, token: str) -> dict:
    try:
        async with httpx.AsyncClient(timeout=30.0) as client:
            res = await client.get(url, headers=_build_headers(token))
    except Exception as e:
        raise HTTPException(status_code=503, detail=f"Could not reach Finnoto API: {e}")

    if res.status_code in (401, 403):
        raise HTTPException(status_code=401, detail="Finnoto token is invalid or expired.")
    if res.status_code == 404:
        raise HTTPException(status_code=404, detail="Not Found")
    if res.status_code >= 400:
        raise HTTPException(
            status_code=502,
            detail=f"Finnoto API error (HTTP {res.status_code}): {res.text[:400]}"
        )
    try:
        return res.json()
    except Exception:
        raise HTTPException(status_code=502, detail="Finnoto returned non-JSON response.")


@router.post("/list")
async def list_reports(request: ReportListRequest, db=Depends(get_db)):
    """
    Returns downloadable business reports with their sheet definitions.
    Tries the Finnoto /api/b/report/list endpoint first, then falls back to
    /api/b/report/search with listing_slug. Every known Finnoto base URL is
    tried (connected base first) so reports are found regardless of which
    environment hosts them.
    """
    base_payload: Dict[str, Any] = {"page": request.page or 1, "limit": request.limit or 50}
    if request.search:
        base_payload["search"] = request.search

    search_payload: Dict[str, Any] = {
        "listing_slug": "ap_downloadable_generate_report",
        "page": request.page or 1,
        "limit": request.limit or 50,
    }
    if request.search:
        search_payload["search"] = request.search

    attempted = []
    auth_error = None

    for base_url in _candidate_bases(db.base_url):
        source = None
        normalised = None

        # Primary path: /api/b/report/list (downloadable reports with sheets)
        try:
            raw_list = await _post(f"{base_url}/api/b/report/list", base_payload, db.token)
            records = _extract_report_records(raw_list)
            if records:
                normalised = _normalise_report_list(records)
                source = "list"
        except HTTPException as exc:
            if exc.status_code in (401, 403):
                auth_error = exc
            source = f"list_failed({exc.status_code})"

        # Fallback: /api/b/report/search with listing_slug (generated reports)
        if not normalised:
            try:
                raw_search = await _post(f"{base_url}/api/b/report/search", search_payload, db.token)
                records = _extract_report_records(raw_search)
                if records:
                    normalised = _normalise_search_results(records)
                    source = "search"
            except HTTPException as exc:
                if exc.status_code in (401, 403):
                    auth_error = exc
                source = f"search_failed({exc.status_code})"

        attempted.append({
            "base_url": base_url,
            "source": source,
            "total": len(normalised) if normalised is not None else 0,
        })

        if normalised:
            print(f"[reports/list] FOUND base_url={base_url} source={source} total={len(normalised)}")
            return {
                "reports": normalised,
                "total": len(normalised),
                "_source": source,
                "_base_url": base_url,
                "_attempted": attempted,
            }

    if auth_error is not None:
        raise auth_error

    print(f"[reports/list] none_found attempted={attempted}")
    return {"reports": [], "total": 0, "_source": None, "_base_url": None, "_attempted": attempted}


def _candidate_bases(connected: str) -> list:
    """Connected base URL first, then all known presets (deduplicated)."""
    ordered = []
    for base in [connected, *FINNOTO_BASE_PRESETS.values()]:
        if base and base not in ordered:
            ordered.append(base)
    return ordered


def _extract_report_records(raw) -> list:
    """Extract a flat list of report records from various Finnoto response shapes."""
    if isinstance(raw, list):
        return raw
    if not isinstance(raw, dict):
        return []
    # Try common envelope keys in priority order
    for key in ("records", "data", "list", "items", "results"):
        val = raw.get(key)
        if isinstance(val, list):
            return val
        if isinstance(val, dict):
            # nested: { data: { records: [...] } }
            for inner in ("records", "list", "items", "data", "results"):
                inner_val = val.get(inner)
                if isinstance(inner_val, list):
                    return inner_val
    return []


def _normalise_report_list(records: list) -> list:
    out = []
    for r in records:
        if not isinstance(r, dict):
            continue
        sheets_raw = r.get("sheets") or r.get("report_sheets") or []
        sheets = []
        for s in sheets_raw:
            if not isinstance(s, dict):
                continue
            sheet_id = s.get("id") or s.get("sheet_id")
            sheet_name = s.get("name") or s.get("sheet_name") or (f"Sheet {sheet_id}" if sheet_id else "Sheet")
            sheets.append({"id": sheet_id, "name": sheet_name})

        report_id = r.get("id")
        name = (r.get("report_name") or r.get("name") or r.get("title") or f"Report {report_id or '?'}")
        out.append({
            "id": report_id,
            "sys_report_id": r.get("sys_report_id"),
            "name": name,
            "description": r.get("report_description") or r.get("description"),
            "sheets": sheets,
            "active": r.get("active", True),
        })
    return out


def _normalise_search_results(records: list) -> list:
    """Normalise generated reports from the search endpoint (no sheets)."""
    out = []
    for r in records:
        if not isinstance(r, dict):
            continue
        name = r.get("report") or r.get("report_name") or r.get("name") or f"Report {r.get('id', '?')}"
        out.append({
            "id": r.get("id"),
            "name": name,
            "url": r.get("url"),
            "status": "ready" if r.get("url") else "processing",
            "description": r.get("filter"),
            "sheets": [],
        })
    return out


@router.post("/debug")
async def debug_reports(db=Depends(get_db)):
    """
    DEBUG: Returns raw responses from both Finnoto endpoints.
    Helps diagnose why reports may not be loading.
    """
    result = {"base_url": db.base_url}

    # Test 1: list endpoint
    try:
        raw = await _post(f"{db.base_url}/api/b/report/list", {"page": 1, "limit": 20}, db.token)
        result["list_status"] = "ok"
        result["list_raw"] = raw
    except HTTPException as exc:
        result["list_status"] = f"error({exc.status_code})"
        result["list_detail"] = exc.detail

    # Test 2: search endpoint
    try:
        raw = await _post(
            f"{db.base_url}/api/b/report/search",
            {"listing_slug": "ap_downloadable_generate_report", "page": 1, "limit": 20},
            db.token,
        )
        result["search_status"] = "ok"
        result["search_raw"] = raw
    except HTTPException as exc:
        result["search_status"] = f"error({exc.status_code})"
        result["search_detail"] = exc.detail

    # Test 3: available-report endpoint (GET)
    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            res = await client.get(
                f"{db.base_url}/api/b/report/available-report",
                headers={**FINNOTO_COMMON_HEADERS, "Authorization": f"Bearer {db.token}"}
            )
        result["available_report_status"] = res.status_code
        if res.status_code == 200:
            result["available_report_raw"] = res.json()
        else:
            result["available_report_body"] = res.text[:400]
    except Exception as e:
        result["available_report_status"] = f"error({str(e)})"

    return result


@router.post("/generate")
async def generate_report(request: ReportGenerateRequest, db=Depends(get_db)):
    """
    Triggers report generation on Finnoto, polls until complete,
    and returns the download URL.
    """
    base_url = db.base_url
    token = db.token

    params: Dict[str, Any] = {}
    if request.date:
        params["date_filter"] = request.date.model_dump(exclude_none=True)

    body: Dict[str, Any] = {"params": params}
    if request.comment:
        body["comment"] = request.comment
    if request.source_type:
        body["source_type"] = request.source_type

    new_report = await _post(
        f"{base_url}/api/b/report/{request.downloadable_report_id}/generate-report",
        body,
        token,
    )

    report_id = new_report.get("id")
    if not report_id:
        raise HTTPException(status_code=502, detail="Finnoto did not return a report ID.")

    for _ in range(30):
        await asyncio.sleep(2)
        status_report = await _get(f"{base_url}/api/b/report/{report_id}", token)
        if status_report.get("processed_at") or status_report.get("url"):
            return {
                "report_id": report_id,
                "url": status_report.get("url"),
                "status": "ready",
            }

    return {
        "report_id": report_id,
        "url": None,
        "status": "processing",
        "message": "Report is still generating. Check back in a moment.",
    }


@router.post("/get-sheet")
async def get_report_sheet(request: ReportSheetRequest, db=Depends(get_db)):
    """
    Fetches computed tabular data for the given sheet_ids of a business report.
    Uses the Finnoto /api/b/report/get-report-sheet JSON endpoint.
    """
    payload: Dict[str, Any] = {
        "business_report_id": request.business_report_id,
        "sheet_ids": request.sheet_ids,
    }
    if request.date:
        payload["date"] = request.date.model_dump(exclude_none=True)
    if request.filters:
        payload["filters"] = request.filters

    try:
        raw = await _post(f"{db.base_url}/api/b/report/get-report-sheet", payload, db.token)
    except HTTPException as exc:
        if exc.status_code in (404, 502):
            raise HTTPException(
                status_code=exc.status_code,
                detail=f"get-report-sheet API not available on this environment ({exc.detail}). Try using the generate flow instead."
            )
        raise

    sheets = _normalise_sheet_response(raw)
    return {"sheets": sheets}


def _normalise_sheet_response(raw) -> list:
    """
    Normalise the get-report-sheet response into a list of
    { id, name, columns, rows } dicts.

    Handles multiple Finnoto response shapes:
    1. { "SheetName": [ {col: val, ...}, ... ] }  — key-per-sheet dict
    2. { data: [ { name, columns, rows } ] }       — wrapped list of sheet objects
    3. [ { name, columns, rows } ]                 — bare list of sheet objects
    4. [ { col: val, ... } ]                       — flat row list (single sheet)
    """
    # ── Case: bare list ──────────────────────────────────────────────────────
    if isinstance(raw, list):
        if not raw:
            return []
        first = raw[0]
        if isinstance(first, dict):
            # Looks like a list of row-dicts (single sheet)
            if "rows" not in first and "columns" not in first and "name" not in first:
                cols = list(first.keys())
                rows = [[r.get(c) for c in cols] for r in raw]
                return [{"id": None, "name": "Sheet 1", "columns": cols, "rows": rows}]
            # Looks like a list of sheet objects
            return _parse_sheet_objects(raw)
        return []

    if not isinstance(raw, dict):
        return []

    # ── Case: wrapped in an envelope key ────────────────────────────────────
    for key in ("data", "records", "sheets", "results", "sheet_data"):
        val = raw.get(key)
        if isinstance(val, list) and val:
            return _normalise_sheet_response(val)   # recurse with unwrapped value

    # ── Case: key-per-sheet dict { "SheetName": [ row, ... ], ... } ─────────
    out = []
    for sheet_name, sheet_rows in raw.items():
        if not isinstance(sheet_rows, list):
            continue
        if not sheet_rows:
            out.append({"id": None, "name": sheet_name, "columns": [], "rows": []})
            continue
        first = sheet_rows[0]
        if isinstance(first, dict):
            cols = list(first.keys())
            rows = [[r.get(c) for c in cols] for r in sheet_rows]
        elif isinstance(first, list):
            # Already row-per-list
            cols = [str(i) for i in range(len(first))]
            rows = sheet_rows
        else:
            cols = []
            rows = sheet_rows
        out.append({"id": None, "name": sheet_name, "columns": cols, "rows": rows})
    return out


def _parse_sheet_objects(sheets_raw: list) -> list:
    """Parse a list of sheet-object dicts into normalised form."""
    normalised = []
    for sheet in sheets_raw:
        if not isinstance(sheet, dict):
            continue
        raw_cols = sheet.get("columns") or sheet.get("headers") or []
        if raw_cols and isinstance(raw_cols[0], dict):
            cols = [c.get("key") or c.get("name") or c.get("label") or str(i) for i, c in enumerate(raw_cols)]
        else:
            cols = [str(c) for c in raw_cols]

        raw_rows = sheet.get("rows") or sheet.get("data") or []
        if raw_rows and isinstance(raw_rows[0], dict):
            rows = [[r.get(c) for c in cols] for r in raw_rows]
        elif raw_rows and isinstance(raw_rows[0], list):
            rows = raw_rows
        else:
            rows = raw_rows

        normalised.append({
            "id": sheet.get("id") or sheet.get("sheet_id"),
            "name": sheet.get("name") or sheet.get("sheet_name") or f"Sheet {sheet.get('id', '')}",
            "columns": cols,
            "rows": rows,
        })
    return normalised
