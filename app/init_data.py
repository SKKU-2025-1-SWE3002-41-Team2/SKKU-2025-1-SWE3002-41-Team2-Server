

def seed_initial_data():
    from app.models import User, ChatSession, Message, ExcelFile
    from app.database import get_db_session

    with get_db_session() as db:
        # 1. User 생성
        user = db.query(User).filter(User.username == "admin").first()
        if not user:
            user = User(username="admin", password="admin123")
            db.add(user)
            db.flush()  # ID 생성 보장
            print("✅ admin 사용자 추가됨")

        # 2. ChatSession 생성
        session = db.query(ChatSession).filter(ChatSession.name == "2025-05-25_기본세션").first()
        if not session:
            session = ChatSession(userId=user.id, name="2025-05-25_기본세션")
            db.add(session)
            db.flush()
            print("✅ 기본 채팅 세션 추가됨")

        # 3. Message 생성
        message = db.query(Message).filter(Message.sessionId == session.id).first()
        if not message:
            db.add(Message(
                sessionId=session.id,
                content="초기 메시지입니다.",
                messageType="nl_command"
            ))
            db.flush()
            print("✅ 기본 메시지 추가됨")

        # 4. ExcelFile 생성
        message = db.query(Message).filter(Message.sessionId == session.id).first()
        file = db.query(ExcelFile).filter(ExcelFile.filename == "샘플.xlsx").first()
        if not file:
            db.add(ExcelFile(
                messageId=message.id,  # ✅ 변경
                filename="샘플.xlsx",
                excelData={}
            ))
            db.flush()
            print("✅ 샘플 엑셀 파일 추가됨")
