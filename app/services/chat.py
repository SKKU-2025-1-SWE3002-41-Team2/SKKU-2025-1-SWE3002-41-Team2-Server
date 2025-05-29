from datetime import datetime

from sqlalchemy.orm import Session
from app.exceptions.http_exceptions import SessionNotFoundException, EmptyMessageAndSheetException, \
    UserNotFoundException
from app.models import ChatSession, Message, ChatSheet, User
from typing import cast, List, Optional, Any

from app.schemas.chat import ChatSessionCreateRequest, MessageRequest
from app.schemas.llm import LLMResponse
from app.services.llm import get_llm_response
from app.services.excel import process_excel_with_commands
from app.utils.timezone import KST

"""
Interface Summary:
- def get_sessions(userId: int, db: Session) -> List[ChatSession]
- def create_session(data: ChatSessionCreateRequest, db: Session) -> ChatSession:
- def delete_session(sessionId: int, db: Session) -> None
- def modify_session(sessionId: int, newName: str, db: Session) -> ChatSession
- def save_message_and_response(sessionId: int, data: MessageRequest, db: Session) -> LLMResponse

Helper Summary:
- def insert_message_to_db(sessionId: int, content: str, senderType: str, db: Session) -> Message
- def upsert_chat_sheet(sessionId: int, sheetData: Optional[Any], db: Session) -> ChatSheet
- def validate_user_exists(userId: int, db: Session) -> None
- def update_session_summary(sessionId: int, summary: str, db: Session) -> None 
- def touch_session(sessionId: int, db: Session)
"""



### read only ###
def get_sessions(userId: int, db: Session) -> List[ChatSession]:
    sessions = (
        db.query(ChatSession)
        .filter(ChatSession.userId == userId)
        .order_by(ChatSession.modifiedAt.desc())
        .all()
    )

    if not sessions:
        raise SessionNotFoundException

    return cast(List[ChatSession], sessions)

def get_messages(session_id: int, db: Session) -> ChatSession:
    session = db.query(ChatSession).filter(ChatSession.id == session_id).first()
    if not session:
        raise SessionNotFoundException
    return session

### modify data ###
def create_session(data: ChatSessionCreateRequest, db: Session) -> LLMResponse:
    if not data.message and not data.sheetData:
        raise EmptyMessageAndSheetException

    validate_user_exists(data.userId, db)

    message_to_ask = None
    sheet_to_ask = None

    session = ChatSession(userId=data.userId, name="New Session")
    db.add(session)
    db.flush()

    if data.message:
        insert_message_to_db(
            sessionId=session.id,
            content=data.message,
            senderType="USER",
            db=db
        )
        message_to_ask = data.message

    # sheet 생성 or 저장
    sheet_to_ask = upsert_chat_sheet(session.id, data.sheetData, db)

    # TODO : llm response
    # llmservice(message_to_ask, sheet_to_ask, session.summary)
    # llm service에서 chat content, xlsx(json), summary
    # db에 저장 후 리턴
    # update_session_summary(session.id, #result.summary, db)

    db.commit()
    return LLMResponse(
        chat="test chat",
        sheetData=[
            ["Name", "Age", "Job"],
            ["Alice", 30, "Engineer"],
            ["Bob", 25, "Designer"]
        ]
    )

def delete_session(sessionId: int, db: Session) -> None:
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if not session:
        raise SessionNotFoundException

    db.delete(session)
    db.commit()

def modify_session(sessionId: int, newName: str, db: Session) -> ChatSession:
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if not session:
        raise SessionNotFoundException

    session.name = newName
    db.commit()
    db.refresh(session)
    return session

def save_message_and_response(sessionId: int, data: MessageRequest, db: Session) -> LLMResponse:
    # 1. 메시지 저장
    message = insert_message_to_db(
        sessionId=sessionId,
        content=data.content,
        senderType="USER",
        db=db
    )

    # 2. 시트 저장 또는 업데이트
    # FIX : sheet 저장을 llm에서 나온 결과물로 저장할 수 있음
    sheet = upsert_chat_sheet(sessionId, data.sheetData, db)

    summary = get_session_summary(sessionId, db)
    # llm service 호출
    # 리턴값은 LLMResultInternal
    result = get_llm_response(
        session_summary=session.summary,
        user_command=message.content,
        excel_bytes=sheet.sheetData
    )

    # 엑셀 파일 수정
    modified_excel_bytes = process_excel_with_commands(
        excel_bytes=sheet.sheetData,
        commands=result.sheetData  # ExcelCommand 리스트
    )

    # 수정된 엑셀 파일을 DB에 저장
    upsert_chat_sheet(session.id, modified_excel_bytes, db)


    db.commit()
    db.refresh(message)

    return LLMResponse(
        chat="test chat",
        sheetData=[
            ["Name", "Age", "Job"],
            ["Alice", 30, "Engineer"],
            ["Bob", 25, "Designer"]
        ]
    )

#### helper ####
def insert_message_to_db(sessionId: int, content: str, senderType: str, db: Session) -> Message:
    message = Message(
        sessionId=sessionId,
        content=content,
        senderType=senderType
    )
    touch_session(sessionId, db)
    db.add(message)
    return message

def upsert_chat_sheet(sessionId: int, sheetData: Optional[Any], db: Session) -> ChatSheet:
    sheet = db.query(ChatSheet).filter(ChatSheet.sessionId == sessionId).first()

    if sheet:
        if sheetData is not None:
            sheet.sheetData = sheetData
        # else: sheetData가 None이면 그대로 유지
    else:
        sheet = ChatSheet(
            sessionId=sessionId,
            sheetData=sheetData if sheetData is not None else []
        )
        db.add(sheet)

    return sheet

def update_session_summary(sessionId: int, summary: str, db: Session) -> None:
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if not session:
        raise SessionNotFoundException

    session.summary = summary

def get_session_summary(sessionId: int, db: Session) -> Optional[str]:
    # 세션의 요약을 가져오는 함수
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if not session:
        raise SessionNotFoundException
    return session.summary

def validate_user_exists(userId: int, db: Session) -> None:
    if not db.query(User).filter(User.id == userId).first():
        raise UserNotFoundException

def touch_session(sessionId: int, db: Session):
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if session:
        session.modifiedAt = datetime.now(KST)
