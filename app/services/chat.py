import base64
import json
from datetime import datetime

from sqlalchemy.orm import Session
from app.exceptions.http_exceptions import SessionNotFoundException, EmptyMessageAndSheetException, \
    UserNotFoundException
from app.models import ChatSession, Message, ChatSheet, User
from typing import cast, List, Optional, Any

from app.schemas.chat import ChatSessionCreateResponse, LLMResponse, MessageResponse
from app.services.excel_service import ExcelService
from app.services.llm_excel_service import LLMExcelService
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
def create_session(userId: int, message: str, sheetData: bytes, db: Session) -> ChatSessionCreateResponse:

    validate_user_exists(userId, db)

    session = ChatSession(userId=userId, name="New Session")
    db.add(session)
    db.flush()
    res = save_message_and_response(session.id, message, sheetData, db)

    return ChatSessionCreateResponse(
        sessionId=session.id,
        sessionName=session.name,
        sheetData =res.sheetData,
        message=res.message
    )

def save_message_and_response(sessionId: int, message: str, sheetData: bytes, db: Session) -> LLMResponse:
    # ✅ 1. 사용자 메시지를 DB에 저장 (USER)
    message = insert_message_to_db(
        sessionId=sessionId,
        content=message,
        senderType="USER",
        db=db
    )

    # ✅ 2. 세션 정보를 DB에서 조회 (없으면 예외 발생)
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if session is None:
        raise SessionNotFoundException

    # ✅ 3. LLM을 호출하여 명령어 해석 및 응답 생성
    # FIXIT: 아래 user_command는 하드코딩되어 있어 나중에 실제 메시지로 대체 필요
    llm_service = LLMExcelService()
    res = llm_service.process_excel_command(
        user_command="A1에서 A10까지 1~10을 차례로 넣고 A1~A10 더한걸 B1에 넣어줘",  # <- message로 바꿔야 함
        summary=session.summary,
        excel_bytes=sheetData
    )

    # ✅ 4. LLM이 생성한 명령어 시퀀스를 바탕으로 엑셀 수정
    excel_service = ExcelService()
    modified_sheet = excel_service.execute_command_sequence(
        excel_bytes=sheetData,
        commands=res.excel_func_sequence
    )

    # ✅ 5. AI의 응답 메시지를 DB에 저장 (AI)
    aiMessage= insert_message_to_db(
        sessionId=sessionId,
        content=res.response,
        senderType="AI",
        db=db
    )

    # ✅ 6. 세션 요약 업데이트
    update_session_summary(
        sessionId=sessionId,
        summary=res.updated_summary,
        db=db
    )

    # ✅ 7. 수정된 엑셀 데이터를 chat_sheet에 업서트
    upsert_chat_sheet(sessionId, modified_sheet, db)

    # ✅ 8. 변경사항 모두 커밋
    db.commit()

    # ✅ 9. 수정된 엑셀 sheet를 base64로 인코딩하여 JSON 응답에 포함
    encoded_sheet = base64.b64encode(modified_sheet).decode('utf-8')

    return LLMResponse(
        sheetData=b"",  # FIXME: encoded_sheet로 바꿔야 정상 작동
        message = MessageResponse(
            id= aiMessage.id,
            content=aiMessage.content,
            createdAt=aiMessage.createdAt,
            senderType=aiMessage.senderType
        )
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



#### helper ####
def insert_message_to_db(sessionId: int, content: str, senderType: str, db: Session) -> Message:
    message = Message(
        sessionId=sessionId,
        content=content,
        senderType=senderType
    )
    touch_session(sessionId, db)
    db.flush()
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
            sheetData=sheetData if sheetData is not None else b""  # 빈 바이트
        )
        db.add(sheet)

    return sheet

def update_session_summary(sessionId: int, summary: str, db: Session) -> None:
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if not session:
        raise SessionNotFoundException

    session.summary = summary

def validate_user_exists(userId: int, db: Session) -> None:
    if not db.query(User).filter(User.id == userId).first():
        raise UserNotFoundException

def touch_session(sessionId: int, db: Session):
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if session:
        session.modifiedAt = datetime.now(KST)
