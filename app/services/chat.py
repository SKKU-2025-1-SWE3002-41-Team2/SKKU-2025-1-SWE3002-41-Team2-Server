import json
from datetime import datetime

from sqlalchemy.orm import Session
from app.exceptions.http_exceptions import SessionNotFoundException, EmptyMessageAndSheetException, \
    UserNotFoundException
from app.models import ChatSession, Message, ChatSheet, User
from typing import cast, List, Optional, Any

from app.schemas.chat import ChatSessionCreateRequest, MessageRequest
from app.schemas.llm import LLMResponse
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
def create_session(data: Any, db: Session) -> LLMResponse:
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

    print("1")
    # sheet 생성 or 저장
    sheet_to_ask = upsert_chat_sheet(session.id, data.sheetData, db)
    print("2")
    # TODO : llm response
    # llmservice(message_to_ask, sheet_to_ask, session.summary)
    # llm service에서 chat content, xlsx(json), summary
    # db에 저장 후 리턴
    # update_session_summary(session.id, #result.summary, db)
    print("3")

    llm_service = LLMExcelService()
    res = llm_service.process_excel_command(
    user_command="A1에서 A10까지 1~10을 차례로 넣고 A1~A10 더한걸 B1에 넣어줘",
    summary=session.summary,
    excel_bytes=sheet_to_ask.sheetData
    )
    print("4")
    excel_service = ExcelService()
    eb = excel_service.convert_json_to_excel_bytes(sheet_to_ask.sheetData)
    modified_bytes = excel_service.execute_command_sequence(
        excel_bytes=eb,
        commands=res.excel_func_sequence
    )
    print("5")
    st = excel_service.load_excel_from_bytes(modified_bytes)
    print(st)
    print("6")
    db.commit()
    return LLMResponse(
        chat="aaa",
        sheetData={
            "sheet_name": "Sheet1",
            "data": {}
          }
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


    # llm service 호출
    #
    # update_session_summary(session.id, #result.summary, db)



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
