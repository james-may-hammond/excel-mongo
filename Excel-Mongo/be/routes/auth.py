"""
Module: auth.py
Description: Authentication route that validates the Finnoto Bearer token
             by making a real trial call to the Finnoto API.
             Supports configurable base URL for dev/staging/production.
"""
import httpx
from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from typing import Optional
from be.db import DEFAULT_FINNOTO_BASE, FINNOTO_COMMON_HEADERS

router = APIRouter()


class LoginRequest(BaseModel):
    token: str
    base_url: Optional[str] = None   # defaults to DEFAULT_FINNOTO_BASE


@router.post("/auth/connect")
async def connect(request: LoginRequest):
    if not request.token or not request.token.strip():
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Missing Finnoto token.")

    token = request.token.strip()
    base_url = (request.base_url or DEFAULT_FINNOTO_BASE).rstrip("/")

    # Validate by hitting the report/list endpoint which exists on all Finnoto environments
    trial_url = f"{base_url}/api/b/report/list"
    trial_payload = {"page": 1, "limit": 1}

    headers = {**FINNOTO_COMMON_HEADERS, "Authorization": f"Bearer {token}"}

    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            res = await client.post(trial_url, json=trial_payload, headers=headers)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=f"Could not reach Finnoto API at {base_url}: {str(e)}"
        )

    if res.status_code in (401, 403):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=(
                "Invalid or expired Finnoto Bearer token. "
                "Please log in to Finnoto and copy a fresh token."
            )
        )

    if res.status_code >= 400:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Finnoto returned an unexpected error (HTTP {res.status_code}): {res.text}"
        )

    return {
        "status": "success",
        "token": token,
        "base_url": base_url,
        "message": "Token validated successfully."
    }