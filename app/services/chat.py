from typing import Any

from fastapi import HTTPException
from sqlalchemy.orm import Session
from app.models import ChatSession
from typing import cast, List

"""
Interface Summary:
- def get_sessions_by_user(userId: int, db: Session) -> List[ChatSession]
- def create_session(userId: int, name: str, db: Session) -> ChatSession
- def delete_session(sessionId: int, db: Session) -> None
"""

def get_sessions_by_user(userId: int, db: Session) -> List[ChatSession]:
    sessions = (
        db.query(ChatSession)
        .filter(ChatSession.userId == userId)
        .order_by(ChatSession.modifiedAt.desc())
        .all()
    )

    if not sessions:
        raise HTTPException(status_code=404, detail="No chat sessions found for this user.")

    return cast(List[ChatSession], sessions)
