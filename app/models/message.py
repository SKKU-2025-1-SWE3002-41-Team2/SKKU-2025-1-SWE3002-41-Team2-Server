from datetime import datetime

from sqlalchemy import Column, Integer, Text, DateTime, ForeignKey, Enum, func
from sqlalchemy.orm import relationship

from app.database import Base
import enum

from app.utils.timezone import KST


class SenderType(enum.Enum):
    USER = "USER"
    AI = "AI"

class Message(Base):
    __tablename__ = "message"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    sessionId = Column(
        Integer,
        ForeignKey("chat_session.id", ondelete="CASCADE"),
        nullable=False
    )
    createdAt = Column(DateTime, default=lambda: datetime.now(KST))
    content = Column(Text, nullable=False)
    senderType = Column(Enum(SenderType), nullable=False)

    session = relationship("ChatSession", back_populates="messages")
