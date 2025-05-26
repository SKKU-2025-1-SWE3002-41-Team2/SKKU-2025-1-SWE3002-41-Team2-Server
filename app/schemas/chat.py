from pydantic import BaseModel, Field
from datetime import datetime
from typing import Optional, Any, List

from app.models.message import SenderType


class ChatSessionResponse(BaseModel):
    sessionId: int = Field(..., alias="id")  # ← 여기!
    userId: int
    name: str | None = None
    modifiedAt: datetime

    class Config:
        from_attributes  = True

class ChatSessionCreateRequest(BaseModel):
    userId: int
    message: Optional[str] = None
    sheetData: Optional[Any] = None

class ChatSessionUpdateRequest(BaseModel):
    name: str  # 수정할 제목


class MessageResponse(BaseModel):
    id: int
    createdAt: datetime
    content: str
    senderType: SenderType

    class Config:
        from_attributes = True

class MessageRequest(BaseModel):
    content: str
    sheetData: Optional[Any] = None



class ChatSessionWithMessagesResponse(BaseModel):
    sessionId: int
    userId: int
    name: str
    modifiedAt: datetime
    sheetData: Optional[Any] = None
    messages: List[MessageResponse]

    class Config:
        from_attributes = True


