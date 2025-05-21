# app/routers/test_db.py
from fastapi import APIRouter, HTTPException
from sqlalchemy.exc import SQLAlchemyError
from pydantic import BaseModel
from typing import List, Optional

from app.database import get_db_session
from app.models.user import User

router = APIRouter()


# 간단한 응답 모델
class UserResponse(BaseModel):
    id: int
    username: str
    created_at: str

    class Config:
        orm_mode = True


# 사용자 생성 요청 모델
class UserCreate(BaseModel):
    username: str
    password: str


# 데이터베이스 상태 확인 엔드포인트
@router.get("/ping")
def ping_database():
    """데이터베이스 연결 상태 확인"""
    try:
        with get_db_session() as db:
            # 간단한 쿼리로 DB 연결 확인
            result = db.execute("SELECT 1").scalar()
            return {"status": "success", "message": "Database connection successful", "result": result}
    except SQLAlchemyError as e:
        raise HTTPException(status_code=500, detail=f"Database error: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")


# 테스트용 사용자 생성 엔드포인트
@router.post("/users/test", response_model=UserResponse)
def create_test_user(user: UserCreate):
    """테스트용 사용자 생성"""
    try:
        with get_db_session() as db:
            # 사용자 중복 확인
            existing = db.query(User).filter(User.username == user.username).first()
            if existing:
                raise HTTPException(status_code=400, detail="Username already exists")

            # 테스트 사용자 생성
            new_user = User(
                username=user.username,
                password=user.password  # 실제로는 해시 처리해야 함
            )
            db.add(new_user)
            db.commit()
            db.refresh(new_user)

            return new_user
    except HTTPException:
        raise
    except SQLAlchemyError as e:
        raise HTTPException(status_code=500, detail=f"Database error: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")


# 사용자 목록 조회 엔드포인트
@router.get("/users/test", response_model=List[UserResponse])
def get_test_users():
    """테스트용 사용자 목록 조회"""
    try:
        with get_db_session() as db:
            users = db.query(User).all()
            return users
    except SQLAlchemyError as e:
        raise HTTPException(status_code=500, detail=f"Database error: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")