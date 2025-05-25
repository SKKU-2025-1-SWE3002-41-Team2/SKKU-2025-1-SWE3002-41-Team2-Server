from sqlalchemy import Column, Integer, String, Text, DateTime, ForeignKey, func
from app.database import Base

class Message(Base):
    __tablename__ = "message"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    sessionId = Column(Integer, ForeignKey("chat_session.id"), nullable=False)
    createdAt = Column(DateTime, server_default=func.now())
    content = Column(Text, nullable=False)
    messageType = Column(String(255), nullable=False)
