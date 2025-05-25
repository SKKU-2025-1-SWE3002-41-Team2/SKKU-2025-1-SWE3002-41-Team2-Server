from sqlalchemy import Column, Integer, Text, DateTime, ForeignKey, Enum, func
from app.database import Base
import enum

class SenderType(enum.Enum):
    USER = "USER"
    AI = "AI"

class Message(Base):
    __tablename__ = "message"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    sessionId = Column(Integer, ForeignKey("chat_session.id"), nullable=False)
    createdAt = Column(DateTime, server_default=func.now())
    content = Column(Text, nullable=False)
    senderType = Column(Enum(SenderType), nullable=False)
