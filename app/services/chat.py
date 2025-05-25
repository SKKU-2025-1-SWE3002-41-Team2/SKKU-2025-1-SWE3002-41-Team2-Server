from sqlalchemy.orm import Session
from app.exceptions.http_exceptions import SessionNotFoundException
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
        raise SessionNotFoundException

    return cast(List[ChatSession], sessions)
