from fastapi import FastAPI
from be.routes.health import router as health_router
from be.routes.insert import router as insertion_router

app = FastAPI(
    title="excel-mongo connector",
    version="1.0.0"
)

app.include_router(health_router)
app.include_router(insertion_router)

@app.get("/")
async def root():
    return {
        "service": "excel-mongo connector",
        "status": "running"
    }

