from os import getenv
from dotenv import load_dotenv
from motor.motor_asyncio import AsyncIOMotorClient

load_dotenv()
MONGO_URL=getenv("MONGO_URL")
# print(MONGO_URL)
if not MONGO_URL: raise RuntimeError("Mongo ENV error.")
mongo_client = AsyncIOMotorClient(MONGO_URL)
db = mongo_client["excel_mongo"]
def get_mongo_client() -> AsyncIOMotorClient:
    return mongo_client