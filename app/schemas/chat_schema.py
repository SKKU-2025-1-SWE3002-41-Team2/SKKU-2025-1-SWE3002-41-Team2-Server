from pydantic import BaseModel, Field
from datetime import datetime
from typing import Optional, Any, List

from app.models.message import SenderType


class ChatSessionResponse(BaseModel):
    """Chat Session 응답 스키마 """
    sessionId: int = Field(..., alias="id")  # ← 여기!
    userId: int
    name: str | None = None
    modifiedAt: datetime

    class Config:
        from_attributes  = True

class MessageResponse(BaseModel):
    """ Message 응답 스키마 """
    id: int
    createdAt: datetime
    content: str
    senderType: SenderType

    class Config:
        from_attributes = True

class ChatSessionCreateResponse(BaseModel):
    """ Chat Session 생성 응답 스키마 """
    sessionId: int
    sessionName: str
    sheetData: Any
    message : MessageResponse

class ChatSessionUpdateRequest(BaseModel):
    """ Chat Session 업데이트 요청 스키마 """
    name: str  # 수정할 제목


class LLMResponse(BaseModel):
    """ message send에 대한 LLM 응답 스키마 """
    sheetData: Any
    message: MessageResponse

    class Config:
        from_attributes = True

class ChatSessionWithMessagesResponse(BaseModel):
    """ Chat Session의 Message 로딩 스키마"""
    sessionId: int
    userId: int
    name: str
    modifiedAt: datetime
    sheetData: Optional[Any] = None
    messages: List[MessageResponse]

    class Config:
        from_attributes = True


