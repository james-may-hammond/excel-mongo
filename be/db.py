import os
from os import getenv
from dotenv import load_dotenv
from motor.motor_asyncio import AsyncIOMotorClient

# Load from root directory or relative directory
load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env"))
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), ".env"))

MONGO_URL=getenv("MONGO_URL")
# print(MONGO_URL)
if not MONGO_URL: raise RuntimeError("Mongo ENV error.")
mongo_client = AsyncIOMotorClient(MONGO_URL)
db = mongo_client["excel_mongo"]
def get_mongo_client() -> AsyncIOMotorClient:
    return mongo_client