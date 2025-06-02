from typing import List, Optional, Dict, Any
import os
from datetime import datetime
from openai import OpenAI
from pydantic import BaseModel, Field

from app.schemas.excel_schemas import LLMExcelResponse, ExcelCommand
from app.services.excel_service import ExcelService
from app.services.excel_commands import CommandType


# Structured Output을 위한 Pydantic 모델
# 이 모델은 OpenAI의 Structured Output 기능을 사용하여
# 명령어 시퀀스를 정의하는 데 사용됩니다.
class ExcelCommandOutput(BaseModel):
    command_type: str = Field(description="명령어 타입")
    target_range: str = Field(description="대상 셀 범위 (예: A1:B10)")
    parameters: Dict[str, Any] = Field(description="명령어 파라미터")


class LLMResponseOutput(BaseModel):
    """LLM 응답 출력 구조"""
    response: str = Field(description="사용자에게 보여줄 한국어 응답")
    commands: List[ExcelCommandOutput] = Field(description="실행할 엑셀 명령어 시퀀스")
    summary: str = Field(description="이번 응답의 내용을 반영한 갱신된 요약")


def get_openai_friendly_schema():
    return {
        "name": "LLMResponseOutput",  # ✅ 이름
        "schema": {                   # ✅ 실제 스키마 내용
            "type": "object",
            "title": "LLMResponseOutput",
            "properties": {
                "response": {
                    "type": "string",
                    "description": "사용자에게 보여줄 한국어 응답"
                },
                "commands": {
                    "type": "array",
                    "description": "실행할 엑셀 명령어 시퀀스",
                    "items": {
                        "type": "object",
                        "properties": {
                            "command_type": {"type": "string", "description": "명령어 타입"},
                            "target_range": {"type": "string", "description": "대상 셀 범위 (예: A1:B10)"},
                            "parameters": {"type": "object", "description": "명령어 파라미터"}
                        },
                        "required": ["command_type", "target_range", "parameters"]
                    }
                },
                "summary": {
                    "type": "string",
                    "description": "이번 응답의 내용을 반영한 갱신된 요약"
                }
            },
            "required": ["response", "commands", "summary"]
        }
    }


class LLMExcelService:
    """LLM과 엑셀 통합 서비스"""

    def __init__(self):
        # 환경변수에서 API 키 가져오기
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY 환경변수가 설정되지 않았습니다.")

        self.client = OpenAI(api_key=api_key)
        self.excel_service = ExcelService()



    def process_excel_command(
            self,
            user_command: str,
            summary: str,
            excel_bytes: bytes
    ) -> LLMExcelResponse:
        """
        사용자 명령을 처리하여 엑셀 명령어 시퀀스 생성

        작동 과정:
        1. 현재 엑셀 파일의 내용을 분석하여 컨텍스트 생성
        2. GPT에게 역할과 사용 가능한 명령어를 설명하는 시스템 프롬프트 생성
        3. 사용자의 명령과 현재 상황을 포함한 프롬프트 생성
        4. GPT-4에 structured output 형식으로 요청
        5. 응답을 파싱하여 ExcelCommand 객체로 변환
        6. 채팅 요약 업데이트
        """

        # 1. 엑셀 파일의 현재 상태를 분석
        excel_context = self._analyze_excel_context(excel_bytes)

        # 2. 시스템 프롬프트 구성
        system_prompt = self._create_system_prompt()

        # 3. 사용자 프롬프트 구성
        user_prompt = self._create_user_prompt(
            summary,
            user_command,
            excel_context
        )
        print("in 1")
        # 4. OpenAI Structured Output 사용
        completion = self.client.beta.chat.completions.parse(
            model="gpt-4.1",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            response_format={"type": "json_schema", "json_schema": get_openai_friendly_schema()},
            max_tokens=1 << 15,
            temperature=0.7 # (온도 조절: 0.7은 적당한 창의성)
        )
        print("in 2")
        print("✅ 프롬프트 토큰:", completion.usage.prompt_tokens)
        print("✅ 응답 토큰:", completion.usage.completion_tokens)
        print("✅ 총 토큰:", completion.usage.total_tokens)

        # 5. 응답 파싱
        parsed_response = completion.choices[0].message.parsed
        print("💬 GPT 원문 응답:", completion.choices[0].message.content)
        print("🔎 Parsed 결과:", parsed_response)

        # 6. ExcelCommand 객체로 변환
        commands = []
        for cmd in parsed_response.commands:
            commands.append(ExcelCommand(
                command_type=cmd.command_type,
                target_range=cmd.target_range,
                parameters=cmd.parameters
            ))

        # 7. 채팅 요약 업데이트


        # 8. 엑셀 명령어 실행
        # 인수인계 파일에서 이전에 설명한 excel_service.execute_command 메서드를 사용하여
        # 각 명령어를 실행합니다.
        return LLMExcelResponse(
            response=parsed_response.response,
            updated_summary=parsed_response.summary or "",
            excel_func_sequence=commands
        )

    def _analyze_excel_context(self, excel_bytes: bytes) -> str:
        """
        엑셀 파일의 현재 상태를 분석하여 GPT가 이해할 수 있는 텍스트로 변환

        분석 내용:
        - 시트의 크기 (행/열 개수)
        - 데이터가 있는 셀의 위치와 내용 샘플
        - 현재 적용된 수식이 있다면 그 정보
        """
        try:
            workbook = self.excel_service.load_excel_from_bytes(excel_bytes)
            ws = workbook.active

            # 데이터가 있는 범위 확인
            max_row = ws.max_row
            max_col = ws.max_column

            # 간단한 요약 생성
            context = f"현재 엑셀 시트: {max_row}행 x {max_col}열\n"

            # 데이터가 있는 셀들의 샘플 수집
            sample_data = []
            formula_cells = []

            for row in range(1, min(11, max_row + 1)):  # 최대 10행까지
                for col in range(1, min(11, max_col + 1)):  # 최대 10열까지
                    cell = ws.cell(row=row, column=col)
                    if cell.value:
                        cell_ref = cell.coordinate

                        # 수식인지 확인
                        if isinstance(cell.value, str) and cell.value.startswith('='):
                            formula_cells.append(f"{cell_ref}: {cell.value}")
                        else:
                            sample_data.append(f"{cell_ref}: {cell.value}")

            if sample_data:
                context += "\n데이터 샘플:\n" + "\n".join(sample_data[:20])

            if formula_cells:
                context += "\n\n수식:\n" + "\n".join(formula_cells)

            return context

        except Exception as e:
            return f"엑셀 파일 분석 중 오류: {str(e)}"

    def _create_system_prompt(self) -> str:
        """GPT의 역할과 사용 가능한 명령어를 정의하는 시스템 프롬프트"""
        return """당신은 엑셀 파일 편집을 도와주는 AI 어시스턴트입니다.
사용자의 자연어 명령을 이해하고, 이를 구체적인 엑셀 명령어 시퀀스로 변환합니다.

사용 가능한 명령어 타입:
- 함수: sum(합계), average(평균), count(개수), max(최대값), min(최소값)
- 서식: bold(굵게), italic(기울임), underline(밑줄), font_color(글자색), fill_color(배경색), border(테두리), font_size(글자크기), font_name(글꼴)
- 데이터: set_value(값 설정), clear(지우기), merge(병합), unmerge(병합 해제)
- 정렬: align_left(왼쪽 정렬), align_center(가운데 정렬), align_right(오른쪽 정렬), align_top(위쪽 정렬), align_middle(중간 정렬), align_bottom(아래쪽 정렬)

명령어 작성 규칙:
1. target_range는 Excel 형식으로 표현 (예: "A1", "B2:C5")
2. 색상은 16진수 6자리로 표현 (예: "FF0000" = 빨간색, "0000FF" = 파란색)
3. 명령어는 실행 순서를 고려하여 논리적으로 배치
4. 수식 명령의 경우 parameters에 'range' 키로 계산 범위 지정
5. summary는 입력받은 summary와 이번 응답에서의 엑셀 시퀀스를 통한 변경점을 반영해 갱신해서 응답
6. summary는 갱신해서 1000자 이하로 응답
7. 모든 명령어는 `parameters` 필드를 반드시 포함해야 합니다.
- 파라미터가 필요한 명령어는 실제 키-값 쌍을 입력합니다.
- 파라미터가 필요 없는 명령어는 다음을 사용해 의미를 명시합니다:
    - `{"note": "no parameters needed"}`
 
예시:
- B2:B10의 합계를 B11에 표시: command_type="sum", target_range="B11", parameters={"range": "B2:B10"}
- A1 셀을 굵게: command_type="bold", target_range="A1", parameters={"note": "no parameters needed"}
- C1:C5를 빨간색으로: command_type="font_color", target_range="C1:C5", parameters={"color": "FF0000"}

응답은 항상 친절하고 명확한 한국어로 작성하세요."""

    def _create_user_prompt(
            self,
            summary: str,
            user_command: str,
            excel_context: str
    ) -> str:
        """사용자의 명령과 현재 상황을 포함한 프롬프트"""
        return f"""이전 대화 요약:


현재 세션 요약:
{summary}

현재 엑셀 파일 상태:
{excel_context}

사용자 명령:
{user_command}

위 정보를 바탕으로 사용자의 명령을 수행하기 위한 엑셀 명령어 시퀀스를 생성하고,
사용자에게 친절한 한국어 응답을 작성해주세요."""
