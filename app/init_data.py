from io import BytesIO

from openpyxl import Workbook


def seed_initial_data():
    from app.models import User, ChatSession, Message, ChatSheet
    from app.database import get_db_session

    db = next(get_db_session())  # ✅ 세션 꺼내기
    try:
        # 1. User 생성
        user = db.query(User).filter(User.username == "admin").first()
        if not user:
            user = User(username="admin", password="admin123")
            db.add(user)
            db.flush()
            print("✅ admin 사용자 추가됨")

        # 2. ChatSession 생성
        session = db.query(ChatSession).filter(ChatSession.name == "2025-06-23_기본세션").first()
        if not session:
            session = ChatSession(userId=user.id, name="2025-06-23_기본세션")
            db.add(session)
            db.flush()
            print("✅ 기본 채팅 세션 추가됨")

        # 3. Message 생성
        message = db.query(Message).filter(Message.sessionId == session.id).first()
        if not message:
            message = Message(
                sessionId=session.id,
                content="초기 메시지입니다.",
                senderType="USER"
            )
            db.add(message)
            print("✅ 기본 메시지 추가됨")

        # 4. ChatSheet 생성
        sheet = db.query(ChatSheet).filter(ChatSheet.sessionId == session.id).first()
        if not sheet:
            sheet = ChatSheet(
                sessionId=session.id,
                sheetData=create_default_sheet_binary()
            )
            db.add(sheet)
            print("✅ 기본 시트 데이터 추가됨")

        db.commit()

    finally:
        db.close()


def create_default_sheet_binary() -> bytes:
    data = [
        ["Item", "Price"],
        ["Notebook", 1200],
        ["Pen", 300]
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for row in data:
        ws.append(row)

    stream = BytesIO()
    wb.save(stream)  # 엑셀 포맷으로 저장
    return stream.getvalue()  # 바이너리 반환e()