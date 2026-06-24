"""
Module: health_test.py
Description: Unit tests for the health check API route.
Dependencies: pytest, httpx
"""
import pytest
from httpx import AsyncClient


@pytest.mark.asyncio
async def test_health():
    async with AsyncClient() as client:
        response = await client.get("http://localhost:8000/health")

    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "ok"