from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from app.database import get_db_session
from app.models import ChatSession, Message
from app.schemas.chat import ChatSessionResponse
from app.services.chat import get_sessions_by_user

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

# @router.post("/messages")
# def create_message(session_id: int, content: str, message_type: str, db: Session = Depends(get_db_session)):
#     message = Message(sessionId=session_id, content=content, messageType=message_type)
#     db.add(message)
#     db.commit()
#     db.refresh(message)
#     return message