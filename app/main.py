from fastapi import FastAPI
from app.database import init_db
from dotenv import load_dotenv

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

