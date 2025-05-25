from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, JSON, func
from app.database import Base

class ExcelFile(Base):
    __tablename__ = "excel_file"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    messageId = Column(Integer, ForeignKey("message.id"), nullable=False)  # ✅ 수정됨
    filename = Column(String(255), nullable=False)
    uploadedAt = Column(DateTime, server_default=func.now())
    lastModified = Column(DateTime)
    excelData = Column(JSON)
