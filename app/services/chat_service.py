import base64
from datetime import datetime

from sqlalchemy.orm import Session
from app.exceptions.http_exceptions import SessionNotFoundException, \
    UserNotFoundException
from app.models import ChatSession, Message, ChatSheet, User
from typing import cast, List, Optional, Any


from app.schemas.chat_schema import ChatSessionCreateResponse, MessageResponse, LLMMessageResponse

from app.services.llm_service import get_llm_response
from app.services.excel_service import process_excel_with_commands
from app.utils.timezone import KST

"""
Interface Summary:
- def get_sessions(userId: int, db: Session) -> List[ChatSession]
- def get_messages(session_id: int, db: Session) -> ChatSession
- def create_session(userId: int, message: str, sheetData: bytes, db: Session) -> ChatSessionCreateResponse
- def save_message_and_response(sessionId: int, message: str, sheetData: bytes, db: Session) -> LLMResponse
- def delete_session(sessionId: int, db: Session) -> None
- def modify_session(sessionId: int, newName: str, db: Session) -> ChatSession

Helper Summary:
- def insert_message_to_db(sessionId: int, content: str, senderType: str, db: Session) -> Message
- def upsert_chat_sheet(sessionId: int, sheetData: Optional[Any], db: Session) -> ChatSheet
- def update_session_summary(sessionId: int, summary: str, db: Session) -> None
- def validate_user_exists(userId: int, db: Session) -> None
- def touch_session(sessionId: int, db: Session) -> None
"""



### read only ###
def get_sessions(userId: int, db: Session) -> List[ChatSession]:
    """
    주어진 사용자 ID에 해당하는 모든 채팅 세션을 최신순으로 조회합니다.

    Args:
        userId (int): 사용자 ID
        db (Session): SQLAlchemy DB 세션

    Returns:
        List[ChatSession]: 채팅 세션 리스트

    Raises:
        SessionNotFoundException: 세션이 존재하지 않을 경우
    """
    sessions = (
        db.query(ChatSession)
        .filter(ChatSession.userId == userId)
        .order_by(ChatSession.modifiedAt.desc())
        .all()
    )

    if not sessions:
        raise SessionNotFoundException()

    return cast(List[ChatSession], sessions)

def get_messages(session_id: int, db: Session) -> ChatSession:
    """
    특정 세션 ID에 해당하는 채팅 세션을 조회하고 메시지를 포함합니다.

    Args:
        session_id (int): 세션 ID
        db (Session): SQLAlchemy DB 세션

    Returns:
        ChatSession: 해당 세션 객체

    Raises:
        SessionNotFoundException: 세션이 존재하지 않을 경우
    """
    session = db.query(ChatSession).filter(ChatSession.id == session_id).first()
    if not session:
        raise SessionNotFoundException()
    return session

### modify data ###
def create_session(userId: int, message: str, sheetData: bytes, db: Session) -> ChatSessionCreateResponse:
    """
    새로운 채팅 세션을 생성하고 첫 사용자 메시지를 저장한 뒤 LLM 응답을 반환합니다.

    Args:
        userId (int): 사용자 ID
        message (str): 사용자 입력 메시지
        sheetData (bytes): 엑셀 시트 데이터
        db (Session): SQLAlchemy DB 세션

    Returns:
        ChatSessionCreateResponse: 생성된 세션 정보 및 초기 응답 데이터
    """
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

def save_message_and_response(sessionId: int, message: str, sheetData: bytes, db: Session) -> LLMMessageResponse:
    """
       세션에 사용자 메시지를 저장하고 LLM으로부터 응답을 받아 처리 및 저장합니다.

       Args:
           sessionId (int): 채팅 세션 ID
           message (str): 사용자 입력 메시지
           sheetData (bytes): 엑셀 시트 데이터
           db (Session): SQLAlchemy DB 세션

       Returns:
           LLMMessageResponse: LLM의 응답 메시지 및 수정된 엑셀 시트 데이터 (Base64 인코딩)

       Raises:
           SessionNotFoundException: 세션이 존재하지 않을 경우
       """
    # 1. 사용자 메시지를 DB에 저장 (USER)
    saved_message = insert_message_to_db(
        sessionId=sessionId,
        content=message,
        senderType="USER",
        db=db
    )

    # 2. 세션 정보를 DB에서 조회 (없으면 예외 발생)
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if session is None:
        raise SessionNotFoundException()

    # 3. LLM을 호출하여 명령어 해석 및 응답 생성
    response_result = get_llm_response(
        #chat_session의 summary를 가져오도록 구현 필요
        session_summary=session.summary,
        user_command=message,
        excel_bytes =sheetData
    )

    # 4. LLM이 생성한 명령어 시퀀스를 바탕으로 엑셀 수정
    modified_excel_bytes = process_excel_with_commands(
        excel_bytes=sheetData,
        commands=response_result.cmd_seq  # ExcelCommand 리스트
    )

    # 5. AI의 응답 메시지를 DB에 저장 (AI)
    ai_message= insert_message_to_db(
        sessionId=sessionId,
        content=response_result.chat,
        senderType="AI",
        db=db
    )

    # 6. 세션 요약 업데이트
    update_session_summary(
        sessionId=sessionId,
        summary=response_result.summary,
        db=db
    )

    # 7. 수정된 엑셀 데이터를 chat_sheet에 업서트
    upsert_chat_sheet(sessionId, modified_excel_bytes, db)

    # 8. 변경사항 모두 커밋
    db.commit()

    # 9. 수정된 엑셀 sheet를 base64로 인코딩하여 JSON 응답에 포함
    encoded_sheet = base64.b64encode(modified_excel_bytes).decode('utf-8')

    return LLMMessageResponse(
        sheetData=encoded_sheet,
        message = MessageResponse(
            id= ai_message.id,
            content=ai_message.content,
            createdAt=ai_message.createdAt,
            senderType=ai_message.senderType
        )
    )


def delete_session(sessionId: int, db: Session) -> None:
    """
    특정 세션 ID에 해당하는 채팅 세션을 삭제합니다.

    Args:
        sessionId (int): 세션 ID
        db (Session): SQLAlchemy DB 세션

    Raises:
        SessionNotFoundException: 세션이 존재하지 않을 경우
    """
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if not session:
        raise SessionNotFoundException()

    db.delete(session)
    db.commit()

def modify_session(sessionId: int, newName: str, db: Session) -> ChatSession:
    """
    특정 세션의 이름을 수정합니다.

    Args:
        sessionId (int): 세션 ID
        newName (str): 새 세션 이름
        db (Session): SQLAlchemy DB 세션

    Returns:
        ChatSession: 이름이 수정된 세션 객체

    Raises:
        SessionNotFoundException: 세션이 존재하지 않을 경우
    """
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if not session:
        raise SessionNotFoundException()

    session.name = newName
    db.commit()
    db.refresh(session)
    return session

#### helper ####
def insert_message_to_db(sessionId: int, content: str, senderType: str, db: Session) -> Message:
    """
    특정 세션에 메시지를 추가로 저장합니다.

    Args:
        sessionId (int): 세션 ID
        content (str): 메시지 내용
        senderType (str): 메시지 발신자 타입 ("USER" 또는 "AI")
        db (Session): SQLAlchemy DB 세션

    Returns:
        Message: 저장된 메시지 객체
    """
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
    """
    세션에 대응하는 ChatSheet 데이터를 삽입하거나 갱신합니다.

    Args:
        sessionId (int): 세션 ID
        sheetData (bytes | None): 엑셀 데이터
        db (Session): SQLAlchemy DB 세션

    Returns:
        ChatSheet: 삽입되거나 갱신된 시트 객체
    """
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
    """
        세션의 summary 필드를 업데이트합니다.

        Args:
            sessionId (int): 세션 ID
            summary (str): 업데이트할 요약 정보
            db (Session): SQLAlchemy DB 세션

        Raises:
            SessionNotFoundException: 세션이 존재하지 않을 경우
        """
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if not session:
        raise SessionNotFoundException()

    session.summary = summary

def validate_user_exists(userId: int, db: Session) -> None:
    """
       해당 사용자 ID가 존재하는지 검증합니다.

       Args:
           userId (int): 사용자 ID
           db (Session): SQLAlchemy DB 세션

       Raises:
           UserNotFoundException: 사용자가 존재하지 않을 경우
       """
    if not db.query(User).filter(User.id == userId).first():
        raise UserNotFoundException()

def touch_session(sessionId: int, db: Session):
    """
       세션의 modifiedAt 필드를 현재 시각으로 갱신합니다.

       Args:
           sessionId (int): 세션 ID
           db (Session): SQLAlchemy DB 세션
    """
    session = db.query(ChatSession).filter(ChatSession.id == sessionId).first()
    if session:
        session.modifiedAt = datetime.now(KST)
