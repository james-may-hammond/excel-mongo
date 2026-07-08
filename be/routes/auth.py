"""
Module: auth.py
Description: Authentication route that validates the Finnoto Bearer token
             by making a real trial call to the Finnoto API.
"""
import httpx
from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel

router = APIRouter()

FINNOTO_TRIAL_URL = "https://eapi.finnoto.dev/api/b/expense/search"
FINNOTO_TRIAL_PAYLOAD = {
    "limit": 1,
    "is_draft": False,
    "listing_slug": "ef_expenses",
    "document_type_identifier": "employee_expenses",
    "ignore_dto_all": True
}

class LoginRequest(BaseModel):
    token: str

@router.post("/auth/connect")
async def connect(request: LoginRequest):
    if not request.token or not request.token.strip():
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Missing Finnoto token.")

    token = request.token.strip()

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Authorization": f"Bearer {token}",
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

    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            res = await client.post(FINNOTO_TRIAL_URL, json=FINNOTO_TRIAL_PAYLOAD, headers=headers)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=f"Could not reach Finnoto API: {str(e)}"
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
            detail=f"Finnoto returned an unexpected error (HTTP {res.status_code}). Try again later."
        )

    return {
        "status": "success",
        "token": token,
        "message": "Token validated successfully."
    }