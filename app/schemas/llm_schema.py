from pydantic import BaseModel
from typing import Any

class ResponseResult(BaseModel):
    chat: str
    cmd_seq: Any
    summary: str