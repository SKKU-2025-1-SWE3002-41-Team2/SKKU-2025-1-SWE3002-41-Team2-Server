from pydantic import BaseModel, Field
from datetime import datetime

class ChatSessionResponse(BaseModel):
    sessionId: int = Field(..., alias="id")  # ← 여기!
    userId: int
    name: str | None = None
    modifiedAt: datetime

    class Config:
        from_attributes  = True
