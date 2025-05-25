from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy.orm import Session
from app.database import get_db_session
from app.schemas.chat import ChatSessionResponse, ChatSessionCreateRequest
from app.schemas.llm import LLMResponse
from app.services.chat import get_sessions_by_user, create_session_or_add_message, create_session_or_add_message

router = APIRouter()

@router.get(
    "/sessions",
    response_model=list[ChatSessionResponse],
    responses={
        404: {"description": "No chat sessions found for this user"},
    },
)
def get_sessions(userId: int = Query(...), db: Session = Depends(get_db_session)):
    return get_sessions_by_user(userId=userId, db=db)


@router.post(
    "/sessions/create",
    response_model=LLMResponse,
    status_code=status.HTTP_201_CREATED,
    summary="Create a session or add message/sheet",
    responses={
        201: {"description": "Session created or message/sheet data added successfully."},
        400: {"description": "Either message or sheetData must be provided."},
    }
)
def create_session(data: ChatSessionCreateRequest, db: Session = Depends(get_db_session)):
    res = create_session_or_add_message(data, db)
    return res



#메시지 모두 불러오기

# 메시지 보내기
# @router.post("/messages")
# def create_message(session_id: int, content: str, message_type: str, db: Session = Depends(get_db_session)):
#     message = Message(sessionId=session_id, content=content, messageType=message_type)
#     db.add(message)
#     db.commit()
#     db.refresh(message)
#     return message

