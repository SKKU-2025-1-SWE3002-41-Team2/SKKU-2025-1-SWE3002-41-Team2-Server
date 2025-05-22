# database.py
import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# 환경변수에서 데이터베이스 URL 가져오기
DATABASE_URL = os.getenv("DATABASE_URL")

# 또는 개별 변수로 URL 구성 (선택사항)
# DB_HOST = os.getenv("DB_HOST", "localhost")
# DB_PORT = os.getenv("DB_PORT", "3307")
# DB_USER = os.getenv("DB_USER", "user")
# DB_PASSWORD = os.getenv("DB_PASSWORD", "excel_pass")
# DB_NAME = os.getenv("DB_NAME", "excel_platform_db")
# DATABASE_URL = f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"

print(f"데이터베이스 연결 URL: {DATABASE_URL}")

# SQLAlchemy 엔진 생성 (로컬 개발용 - 간단한 설정)
engine = create_engine(
    DATABASE_URL,
    echo=False,  # SQL 로그 출력 끄기 (콘솔 깔끔하게)
    pool_pre_ping=True,  # 연결 상태 확인
    # 로컬 개발용 - 최소한의 설정만
    connect_args={
        "charset": "utf8mb4"
    }
)

# 세션 팩토리 생성
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# Base 클래스 생성
Base = declarative_base()

# 데이터베이스 세션 의존성 함수
def get_db():
    """
    FastAPI 의존성 주입용 데이터베이스 세션 생성 함수
    """
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# 연결 테스트 함수 (간단한 버전)
def test_connection():
    """
    데이터베이스 연결 테스트 (로컬 개발용)
    """
    try:
        from sqlalchemy import text
        with engine.connect() as connection:
            # 간단한 연결 테스트
            result = connection.execute(text("SELECT 1"))
            print("✅ 데이터베이스 연결 성공!")
            return True
    except ImportError as e:
        print(f"❌ 패키지 설치 필요: pip install pymysql sqlalchemy python-dotenv cryptography")
        return False
    except Exception as e:
        print(f"❌ 데이터베이스 연결 실패: {e}")
        print("💡 해결 방법: docker-compose down && docker-compose up -d")
        return False

if __name__ == "__main__":
    # 직접 실행 시 연결 테스트
    test_connection()