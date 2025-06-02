from typing import List, Any, Optional, Dict

from pydantic import BaseModel


class CommandSequenceTestRequest(BaseModel):
    """명령어 시퀀스 테스트 요청 스키마"""
    commands: List[Dict[str, Any]]  # 실행할 명령어 시퀀스
    initial_data: Optional[List[List[Any]]] = None  # 초기 데이터 (선택사항)

class CommandSequenceTestResponse(BaseModel):
    """명령어 시퀀스 테스트 응답 스키마"""
    success: bool  # 성공 여부
    message: str  # 결과 메시지
    initial_data: List[List[Any]]  # 초기 데이터
    final_data: List[List[Any]]  # 최종 데이터
    executed_commands: List[Dict[str, Any]]  # 실행된 명령어들
    errors: List[str]  # 오류 메시지들
