from pydantic import BaseModel, Field
from datetime import datetime
from typing import Optional, Any

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