import pytest
from httpx import AsyncClient

@pytest.mark.asyncio
async def test_collections():
    async with AsyncClient() as client:
        response = await client.get("http://localhost:8000/collections")

        assert response.status_code == 200
        data = response.json()
        assert "collections" in data
        