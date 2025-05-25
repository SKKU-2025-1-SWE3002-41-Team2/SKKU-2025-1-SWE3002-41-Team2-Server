from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from app.database import get_db_session
from app.models import ChatSession, Message

router = APIRouter()

@router.get("/sessions")
def get_sessions(db: Session = Depends(get_db_session)):
    return db.query(ChatSession).all()

@router.post("/messages")
def create_message(session_id: int, content: str, message_type: str, db: Session = Depends(get_db_session)):
    message = Message(sessionId=session_id, content=content, messageType=message_type)
    db.add(message)
    db.commit()
    db.refresh(message)
    return message