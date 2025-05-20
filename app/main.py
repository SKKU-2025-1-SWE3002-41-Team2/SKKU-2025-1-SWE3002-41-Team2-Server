# app/main.py
import os
from app import create_app
from dotenv import load_dotenv

# .env 파일 로드 (환경 변수 설정)
load_dotenv()

# 앱 생성
app = create_app()

# 애플리케이션 실행 (직접 실행될 때)
if __name__ == '__main__':
    # 환경 변수에서 포트를 가져오거나 기본값 5000 사용
    port = int(os.environ.get('PORT', 5000))
    # 개발 환경에서는 디버그 모드 활성화
    debug = os.environ.get('FLASK_ENV', 'development') == 'development'

    # 앱 실행
    app.run(host='0.0.0.0', port=port, debug=debug)