# app/database.py
import os
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, scoped_session
from contextlib import contextmanager

# 데이터베이스 URL 설정 (환경 변수에서 가져옴)
DATABASE_URL = os.getenv("DATABASE_URL", "mysql+pymysql://user:password@localhost:3306/excel_platform")

# SQLAlchemy 엔진 생성
engine = create_engine(
    DATABASE_URL,
    pool_size=10,  # 커넥션 풀 사이즈
    max_overflow=20,  # 최대 초과 커넥션
    pool_recycle=3600,  # 커넥션 재활용 시간 (1시간)
    pool_pre_ping=True  # 사용 전 핑 테스트 (연결 유효성 확인)
)

# 세션 팩토리 생성
SessionFactory = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# 스레드 로컬 세션 생성
SessionLocal = scoped_session(SessionFactory)

# Base 클래스 생성 (모델 클래스의 기본 클래스)
Base = declarative_base()
Base.query = SessionLocal.query_property()


@contextmanager
def get_db_session():
    """
    데이터베이스 세션을 제공하는 컨텍스트 매니저

    사용 예:
    with get_db_session() as session:
        users = session.query(User).all()
    """
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception as e:
        session.rollback()
        raise e
    finally:
        session.close()


def init_db():
    """데이터베이스 테이블 초기화"""
    Base.metadata.create_all(bind=engine)


def drop_db():
    """데이터베이스 테이블 삭제 (테스트용)"""
    Base.metadata.drop_all(bind=engine)