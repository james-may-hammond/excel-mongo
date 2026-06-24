"""
Module: update_test.py
Description: Unit tests for the single document update API route.
Dependencies: pytest, httpx
"""
import pytest
from httpx import AsyncClient

@pytest.mark.asyncio
async def test_update():
    insert_data = {
        "collection": "employees",
        "data": {
            "name": "test_update_user",
            "city": "Mumbai",
            "salary": 6000000
        }
    }

    async with AsyncClient() as client:
        insert_response = await client.post("http://localhost:8000/insert", json=insert_data)
        assert insert_response.status_code == 201
        
        fetch_response = await client.post("http://localhost:8000/fetch", json={
            "collection": "employees",
            "filters": {"name": "test_update_user"}
        })
        assert fetch_response.status_code == 200
        fetch_data = fetch_response.json()
        assert fetch_data["count"] >= 1
        record_id = fetch_data["data"][0]["_id"]

        update_data = {
            "collection": "employees",
            "id": record_id,
            "data": {
                "name": "test_update_user",
                "city": "Pune",
                "salary": 6500000
            }
        }
        update_response = await client.post("http://localhost:8000/update", json=update_data)
        assert update_response.status_code == 200
        assert update_response.json()["status"] == "success"

        verify_response = await client.post("http://localhost:8000/fetch", json={
            "collection": "employees",
            "filters": {"name": "test_update_user"}
        })
        assert verify_response.status_code == 200
        verify_data = verify_response.json()
        updated_record = verify_data["data"][0]
        assert updated_record["city"] == "Pune"
        assert updated_record["salary"] == 6500000
