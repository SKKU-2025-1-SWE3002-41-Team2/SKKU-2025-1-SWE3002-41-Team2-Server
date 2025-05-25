from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, func
from app.database import Base

class ChatSession(Base):
    __tablename__ = "chat_session"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    userId = Column(Integer, ForeignKey("user.id"), nullable=False)
    name = Column(String(255))
    summary = Column(String(500),nullable=True)
    createdAt = Column(DateTime, server_default=func.now())
    modifiedAt = Column(DateTime, server_default=func.now())