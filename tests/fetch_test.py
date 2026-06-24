"""
Module: fetch_test.py
Description: Unit tests for the document fetching API route.
Dependencies: pytest, httpx
"""
import pytest
from httpx import AsyncClient

@pytest.mark.asyncio
async def test_fetch():
    async with AsyncClient() as client:
        test_filter = {
            "collection": "employees",
            "filters": {}
        }
        response = await client.post("http://localhost:8000/fetch",json=test_filter)
    print(response.text)
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "success"
    assert "count" in data
    assert "data" in data