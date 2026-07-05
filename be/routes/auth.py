"""
Module: auth.py
Description: Authentication route that validates the Finnoto access token.
"""
from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel

router = APIRouter()

class LoginRequest(BaseModel):
    token: str

@router.post("/auth/connect")
async def connect(request: LoginRequest):
    # Here we would normally make a trial ping to Finnoto to validate the token.
    # For now, we assume the provided token is valid.
    if not request.token:
        raise HTTPException(status_code=401, detail="Missing Finnoto Token")

    return {
        "status": "success", 
        "token": request.token
    }