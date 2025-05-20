from fastapi import FastAPI
# from app.routers import auth, chat, excel  # 기존 라우터 구조 재활용
from app.database import init_db
from dotenv import load_dotenv

load_dotenv()
import os
print("✅ DATABASE_URL =", os.getenv("DATABASE_URL"))
app = FastAPI(title="Excel-LLM Platform")

# DB 초기화 (옵션)
init_db()

# # 라우터 등록
# app.include_router(auth.router, prefix="/api/auth")
# app.include_router(chat.router, prefix="/api/chat")
# app.include_router(excel.router, prefix="/api/excel")

@app.get("/health")
def health_check():
    return {"status": "ok"}
