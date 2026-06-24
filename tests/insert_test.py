"""
Module: insert_test.py
Description: Unit tests for the single document insertion API route.
Dependencies: pytest, httpx
"""
import pytest
from httpx import AsyncClient

@pytest.mark.asyncio
async def test_insert():
    test_data = {
        "collection": "employees",
        "data": {
            "name": "test_user",
            "city": "Mumbai",
            "salary": 5000000
        }
    }

    async with AsyncClient() as client:
        response = await client.post("http://localhost:8000/insert",json=test_data)
        assert response.status_code == 201
        data = response.json()
        assert data["status"] == "sucess"