from pydantic import BaseModel
from typing import Any

class LLMResponse(BaseModel):
    chat: str
    sheetData: Any

    class Config:
        from_attributes = True
