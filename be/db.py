from fastapi import Header, HTTPException
from motor.motor_asyncio import AsyncIOMotorClient
from motor.core import AgnosticDatabase

# created client cache so that a new connection is not opened unnesesarily
client_cache = {}

async def get_db(x_mongo_uri: str = Header(None), x_mongo_db: str = Header(None)) -> AgnosticDatabase:
    print("DEBUG HEADERS - URI:", x_mongo_uri, "DB:", x_mongo_db)
    if not x_mongo_uri or not x_mongo_db:
        raise HTTPException(status_code=401, detail="Missing X-Mongo-URI or X-Mongo-DB headers. Please login in the Add-in.")
    
    if x_mongo_uri not in client_cache:
        try:
            client_cache[x_mongo_uri] = AsyncIOMotorClient(x_mongo_uri, serverSelectionTimeoutMS=5000)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Failed to initialize MongoDB client: {str(e)}")
            
    client = client_cache[x_mongo_uri]
    return client[x_mongo_db]