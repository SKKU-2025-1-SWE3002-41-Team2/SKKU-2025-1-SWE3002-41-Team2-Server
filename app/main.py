from dotenv import load_dotenv
from .database import init_db
from fastapi import FastAPI

load_dotenv()
import os
app = FastAPI(title="Excel-LLM Platform")

# DB 초기화 (옵션)
init_db()
print("✅ DATABASE_URL =", os.getenv("DATABASE_URL"))
# # 라우터 등록

from app.routers import router as api_router
app.include_router(api_router, prefix="/api")
@app.get("/health")
def health_check():
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)