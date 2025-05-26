from sqlalchemy import Column, Integer, ForeignKey, DateTime, JSON, func
from sqlalchemy.orm import relationship

from app.database import Base

class ChatSheet(Base):
    __tablename__ = "chat_sheet"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    sessionId = Column(
        Integer,
        ForeignKey("chat_session.id", ondelete="CASCADE"),
        nullable=False
    )
    sheetData = Column(JSON, nullable=False)

    session = relationship("ChatSession", back_populates="sheet", passive_deletes=True)