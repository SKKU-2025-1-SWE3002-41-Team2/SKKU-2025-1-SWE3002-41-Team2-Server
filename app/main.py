from dotenv import load_dotenv
from .database import init_db
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware




load_dotenv()
import os
app = FastAPI(title="Excel-LLM Platform")


app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],  # 프론트엔드 주소 (필요시 로 허용)
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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