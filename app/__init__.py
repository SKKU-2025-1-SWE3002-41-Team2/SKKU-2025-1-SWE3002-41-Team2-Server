# app/__init__.py
from flask import Flask
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
from config import Config

# 전역 변수로 데이터베이스 객체 선언
db = SQLAlchemy()


def create_app(config_class=Config):
    """
    애플리케이션 팩토리 함수

    Args:
        config_class: 설정 클래스 (기본값: Config)

    Returns:
        생성된 Flask 애플리케이션
    """
    # Flask 애플리케이션 생성
    app = Flask(__name__)
    app.config.from_object(config_class)

    # CORS 설정 - 프론트엔드와 통신 허용
    CORS(app)

    # 데이터베이스 초기화
#    db.init_app(app)

    # 블루프린트 등록 (나중에 추가될 라우트 모듈)
#    from app.routes import auth_bp, chat_bp, excel_bp
#    app.register_blueprint(auth_bp, url_prefix='/api/auth')
#    app.register_blueprint(chat_bp, url_prefix='/api/chat')
#    app.register_blueprint(excel_bp, url_prefix='/api/excel')

    # 데이터베이스 생성
#    with app.app_context():
        #db.create_all()

#    @app.route('/health')
    def health_check():
        """서버 상태 체크용 엔드포인트"""
        return {'status': 'healthy'}

    return app