import pytest
from httpx import AsyncClient

@pytest.mark.asyncio
async def test_update():
    # Insert new record using /insert first
    insert_data = {
        "collection": "employees",
        "data": {
            "name": "test_update_user",
            "city": "Mumbai",
            "salary": 6000000
        }
    }

    async with AsyncClient() as client:
        # 1. Insert record
        insert_response = await client.post("http://localhost:8000/insert", json=insert_data)
        assert insert_response.status_code == 201
        
        # 2. Fetch record to get its ID
        fetch_response = await client.post("http://localhost:8000/fetch", json={
            "collection": "employees",
            "filters": {"name": "test_update_user"}
        })
        assert fetch_response.status_code == 200
        fetch_data = fetch_response.json()
        assert fetch_data["count"] >= 1
        record_id = fetch_data["data"][0]["_id"]

        # 3. Update record using /update
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

        # 4. Fetch again to verify update
        verify_response = await client.post("http://localhost:8000/fetch", json={
            "collection": "employees",
            "filters": {"name": "test_update_user"}
        })
        assert verify_response.status_code == 200
        verify_data = verify_response.json()
        updated_record = verify_data["data"][0]
        assert updated_record["city"] == "Pune"
        assert updated_record["salary"] == 6500000
