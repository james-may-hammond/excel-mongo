from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from be.routes.health import router as health_router
from be.routes.insert import router as insertion_router
from be.routes.collections import router as collections_router
from be.routes.fetch import router as fetch_router
from be.routes.update import router as update_router
from be.routes.schema import router as schema_router

app = FastAPI(
    title="excel-mongo connector",
    version="1.0.0"
)

# Enable CORS for Office Add-in environments (including Excel Online)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(health_router)
app.include_router(insertion_router)
app.include_router(collections_router)
app.include_router(fetch_router)
app.include_router(update_router)
app.include_router(schema_router)
# Serve the static files for the Excel Task Pane GUI
app.mount("/fe", StaticFiles(directory="fe"), name="fe")

@app.get("/")
async def root():
    return {
        "service": "excel-mongo connector",
        "status": "running"
    }

