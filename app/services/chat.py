from sqlalchemy.orm import Session
from app.exceptions.http_exceptions import SessionNotFoundException, EmptyMessageAndSheetException, \
    UserNotFoundException
from app.models import ChatSession, Message, ChatSheet, User
from typing import cast, List

from app.schemas.chat import ChatSessionCreateRequest
from app.schemas.llm import LLMResponse

"""
Interface Summary:
- def get_sessions_by_user(userId: int, db: Session) -> List[ChatSession]
- def create_session_or_add_message(data: ChatSessionCreateRequest, db: Session) -> ChatSession:
- def delete_session(sessionId: int, db: Session) -> None
"""
"""
todo 
채팅방 이름 수정
채팅방 삭제 기능(포함된 채팅까지 전부삭제)

채팅 히스토리 로딩 
"""



def get_sessions_by_user(userId: int, db: Session) -> List[ChatSession]:
    sessions = (
        db.query(ChatSession)
        .filter(ChatSession.userId == userId)
        .order_by(ChatSession.modifiedAt.desc())
        .all()
    )

    if not sessions:
        raise SessionNotFoundException

    return cast(List[ChatSession], sessions)

def create_session_or_add_message(data: ChatSessionCreateRequest, db: Session) -> LLMResponse:
    if not data.message and not data.sheetData:
        raise EmptyMessageAndSheetException

    validate_user_exists(data.userId, db)

    message_to_ask = None
    sheet_to_ask = None


    session = ChatSession(userId=data.userId, name="New Session")
    db.add(session)
    db.flush()

    if data.message:
        add_message(
            session_id=session.id,
            content=data.message,
            sender_type="USER",
            db=db
        )
        message_to_ask = data.message

    existing_sheet = db.query(ChatSheet).filter(ChatSheet.sessionId == session.id).first()

    sheet_data = data.sheetData or []

    if existing_sheet:
        existing_sheet.sheetData = sheet_data
        sheet_to_ask = existing_sheet
    else:
        sheet = ChatSheet(
            sessionId=session.id,
            sheetData=sheet_data
        )
        db.add(sheet)
        sheet_to_ask = sheet




    # TODO : llm response
    # llm service에서 text와 xlsx를 받아서
    # db에 저장 후 리턴

    db.commit()
    return LLMResponse(
        chat="test chat",
        sheetData=[
            ["Name", "Age", "Job"],
            ["Alice", 30, "Engineer"],
            ["Bob", 25, "Designer"]
        ]
    )

def add_message(session_id: int, content: str, sender_type: str, db: Session) -> Message:
    message = Message(
        sessionId=session_id,
        content=content,
        senderType=sender_type
    )
    db.add(message)
    return message

def validate_user_exists(user_id: int, db: Session) -> None:
    if not db.query(User).filter(User.id == user_id).first():
        raise UserNotFoundException