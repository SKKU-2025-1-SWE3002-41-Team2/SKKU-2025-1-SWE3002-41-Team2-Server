# main.py
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.routers import test_db  # 테스트 라우터 임포트
from dotenv import load_dotenv

load_dotenv()
import os
print("✅ DATABASE_URL =", os.getenv("DATABASE_URL"))
app = FastAPI(title="Excel-LLM Platform")

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 개발용, 실제 환경에서는 제한해야 함
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# DB 초기화
init_db()

# 테스트 라우터 등록
app.include_router(test_db.router, prefix="/api/test", tags=["Database Test"])


# User 정보 하드 코딩
@app.on_event("startup")
def create_admin_user():
    # 여기서 직접 임포트 - 순환 참조 방지
    from app.models.user import User
    from app.database import get_db_session
    # get_db_session 컨텍스트 매니저 사용
    with get_db_session() as db:
        admin_user = db.query(User).filter_by(username="admin").first()

        if not admin_user:
            print("✅ Creating admin user")
            default_user = User(
                username="admin",
                password="admin"
            )
            db.add(default_user)
            db.commit()
            print("✅ Admin user created")
        else:
            print("✅ Admin user already exists")

@app.get("/health")
def health_check():
    return {"status": "ok"}