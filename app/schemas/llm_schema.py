from pydantic import BaseModel
from typing import Any

class LLMResultInternal(BaseModel):
    chat: str
    cmd_seq: Any
    summary: str

class ResponseResult(BaseModel):
    chat: str
    cmd_seq: Any
    summary: str