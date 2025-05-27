from pydantic import BaseModel
from typing import List, Optional, Dict, Any
from datetime import datetime

class ExcelEditRequest(BaseModel):
    """엑셀 편집 요청 스키마"""
    history_summary: str  # 채팅 히스토리 요약
    user_command: str     # 사용자 명령어
    excel_file: bytes     # xlsx 파일 데이터
    chat_session_id: int  # 채팅 세션 ID

class ExcelCommand(BaseModel):
    """엑셀 명령어 스키마"""
    command_type: str     # 명령어 타입 (function, format, data)
    target_range: str     # 대상 셀 범위 (예: A1:B10)
    parameters: Dict[str,]  # 명령어 파라미터

class LLMExcelResponse(BaseModel):
    """LLM 응답 스키마"""
    response: str                           # 사용자에게 보여줄 응답
    updated_summary: str                    # 업데이트된 채팅 요약 (1000자 이하)
    excel_func_sequence: List[ExcelCommand]  # 실행할 엑셀 명령어 시퀀스

class ExcelEditResponse(BaseModel):
    """엑셀 편집 응답 스키마"""
    success: bool
    message: str
    excel_file: Optional[bytes]  # 편집된 xlsx 파일
    ai_response: str              # AI 응답
    updated_summary: str          # 업데이트된 채팅 요약