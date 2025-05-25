from sqlalchemy import Column, Integer, ForeignKey, DateTime, JSON, func
from app.database import Base

class ChatSheet(Base):
    __tablename__ = "chat_sheet"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    sessionId = Column(Integer, ForeignKey("chat_session.id"), nullable=False)
    sheetData = Column(JSON, nullable=False)