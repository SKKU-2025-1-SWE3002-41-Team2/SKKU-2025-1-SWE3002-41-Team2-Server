from datetime import datetime, timezone, timedelta

from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, func
from sqlalchemy.orm import relationship

from app.database import Base
from app.utils.timezone import KST


class ChatSession(Base):
    __tablename__ = "chat_session"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    userId = Column(Integer, ForeignKey("user.id"), nullable=False)
    name = Column(String(255))
    summary = Column(String(500),nullable=True)
    createdAt = Column(DateTime, default=lambda: datetime.now(KST))
    modifiedAt = Column(DateTime, default=lambda: datetime.now(KST))

    messages = relationship("Message", back_populates="session", cascade="all, delete-orphan")
    sheet = relationship("ChatSheet", back_populates="session", uselist=False, cascade="all, delete-orphan")