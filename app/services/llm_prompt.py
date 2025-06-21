# app/services/llm_prompt.py
"""
LLM 프롬프트 템플릿 정의
이 파일은 LLM과의 상호작용에서 사용되는 모든 프롬프트를 관리합니다.
"""

# 시스템 프롬프트 - GPT의 역할과 사용 가능한 명령어를 정의
SYSTEM_PROMPT = """당신은 엑셀 파일 편집을 도와주는 AI 어시스턴트입니다.
사용자의 자연어 명령을 이해하고, 이를 구체적인 엑셀 명령어 시퀀스로 변환합니다.

사용 가능한 명령어 타입 (command_type에 사용할 수 있는 값):
- 함수: sum(합계), average(평균), count(개수), max(최대값), min(최소값)
- 논리 함수: if(조건에 따라 다른 값을 반환), and(모든 조건이 참일 때 참 반환), or(조건 중 하나라도 참이면 참 반환)
- 서식: bold(굵게), italic(기울임), underline(밑줄), font_color(글자색), fill_color(배경색), border(테두리), font_size(글자크기), font_name(글꼴)
- 데이터: set_value(값 설정), clear(지우기), merge(병합), unmerge(병합 해제)
- 정렬: align_left(왼쪽 정렬), align_center(가운데 정렬), align_right(오른쪽 정렬), align_top(위쪽 정렬), align_middle(중간 정렬), align_bottom(아래쪽 정렬)
- 검색: vlookup(세로 방향으로 값을 찾아 해당 데이터를 반환),
hlookup(가로 방향으로 값을 찾아 해당 데이터를 반환),
index(지정된 행과 열의 교차점에 있는 값을 반환),
match(찾는 값이 몇 번째 위치에 있는지 반환)


명령어 작성 규칙:
1. command_type은 위에 나열된 값 중 하나여야 합니다 (소문자로 작성)
2. target_range는 Excel 형식으로 표현 (예: "A1", "B2:C5")
3. 명령어는 실행 순서를 고려하여 논리적으로 배치
4. 수식 명령의 경우 parameters에 계산에 필요한 값들을 배열로 지정
5. summary는 입력받은 summary와 이번 응답에서의 엑셀 시퀀스를 통한 변경점을 반영해 갱신해서 1000자 이하로 응답
6. 엑셀 수식 함수에 대해서는 항상 소숫점을 최대 세째 자리까지 반올림하여 표시합니다.
7. 모든 명령어는 `parameters` 필드를 반드시 포함해야 합니다.
   - 파라미터가 필요한 명령어는 실제 값들을 배열로 입력합니다.
   - 파라미터가 필요 없는 명령어는 빈 배열 []을 사용합니다.

예시:
- B2:B10의 합계를 B11에 표시: command_type="sum", target_range="B11", parameters=["B2:B10"]
- 값 설정: command_type="set_value", target_range="A1", parameters=["Hello"]

- IF 함수: B1 값이 60 이상이면 "합격", 아니면 "불합격"을 A1에 설정
{"command_type": "if", "target_range": "A1", "parameters": ["B1>=60", "합격", "불합격"]}

- AND 함수: B1>50 그리고 C1<100 모두 참일 때 TRUE 반환 (A2 셀)
{"command_type": "and", "target_range": "A2", "parameters": ["B1>50", "C1<100"]}

- VLOOKUP 함수: E1 값을 A2:B10 범위에서 찾아 B 열 값 반환 → G1에 표시
{"command_type": "vlookup", "target_range": "G1", "parameters": ["E1", "A2:B10", 2, False]}

중요: 
- command_type은 반드시 enum에 정의된 값 중 하나여야 합니다
- parameters는 항상 배열(리스트) 형태여야 합니다
- 수식 함수의 경우 parameters[0]에 범위를 넣습니다
- 값 설정의 경우 parameters[0]에 설정할 값을 넣습니다
- 이미 값이 있는 셀의 경우 목적 없이 set_value로 값을 변경하지 않습니다.
- response 필드는 반드시 마크다운(Markdown) 형식으로 작성해야 하며, 표, 코드블록, 강조 등 마크다운 문법을 적극적으로 활용하세요.

응답은 항상 친절하고 명확한 한국어로 작성하세요."""

# 사용자 프롬프트 템플릿
USER_PROMPT_TEMPLATE = """이전 대화 요약:
{summary}

현재 엑셀 파일 상태:
{excel_context}

사용자 명령:
{user_command}

위 정보를 바탕으로 사용자의 명령을 수행하기 위한 엑셀 명령어 시퀀스를 생성하고,
사용자에게 친절한 한국어 응답을 작성해주세요.

반드시 다음 JSON 스키마 형식으로 응답해주세요:
{{
    "response": "사용자에게 보여줄 한국어 응답",
    "commands": [
        {{
            "command_type": "명령어 타입",
            "target_range": "대상 셀 범위",
            "parameters": ["파라미터 값들의 배열"]
        }}
    ],
    "summary": "갱신된 요약 (1000자 이하)"
}}"""

# 엑셀 분석 결과 포맷 템플릿
EXCEL_CONTEXT_TEMPLATE = """현재 엑셀 시트: {rows}행 x {cols}열

데이터 샘플:
{sample_data}

수식:
{formula_data}"""

# 에러 상황에 대한 프롬프트
ERROR_PROMPT = """사용자의 요청을 처리하는 중 문제가 발생했습니다.
명령을 더 구체적으로 설명해주시거나, 다시 시도해주세요."""

# GPT API 응답 스키마
RESPONSE_SCHEMA = {
    "type": "json_schema",
    "json_schema": {
        "name": "LLMResponseOutput",
        "strict": True,  # Structured Outputs 활성화
        "schema": {
            "type": "object",
            "properties": {
                "response": {
                    "type": "string",
                    "description": "사용자에게 보여줄 한국어 응답. 마크다운 형식으로 생성"
                },
                "commands": {
                    "type": "array",
                    "description": "실행할 엑셀 명령어 시퀀스",
                    "items": {
                        "type": "object",
                        "properties": {
                            "command_type": {
                                "type": "string",
                                "description": "명령어 타입",
                                "enum": [
                                    # 함수 관련
                                    "sum", "average", "count", "max", "min",
                                    # 데이터 관련
                                    "set_value", "clear", "merge", "unmerge"
                                ]
                            },
                            "target_range": {
                                "type": "string",
                                "description": "대상 셀 범위 (예: A1:B10)"
                            },
                            "parameters": {
                                "type": "array",
                                "description": "명령어 파라미터 배열",
                                "items": {
                                    "type": ["string", "number", "boolean", "null"]
                                }
                            }
                        },
                        "required": ["command_type", "target_range", "parameters"],
                        "additionalProperties": False
                    }
                },
                "summary": {
                    "type": "string",
                    "description": "이번 응답의 내용을 반영한 갱신된 요약"
                }
            },
            "required": ["response", "commands", "summary"],
            "additionalProperties": False
        }
    }
}


def create_user_prompt(summary: str, user_command: str, excel_context: str) -> str:
    """
    사용자 프롬프트를 생성합니다.

    Args:
        summary: 이전 대화 요약
        user_command: 사용자의 명령
        excel_context: 현재 엑셀 파일 상태

    Returns:
        완성된 사용자 프롬프트
    """
    return USER_PROMPT_TEMPLATE.format(
        summary=summary or "없음",
        excel_context=excel_context,
        user_command=user_command
    )


def create_excel_context(rows: int, cols: int, sample_data: list, formula_data: list) -> str:
    """
    엑셀 파일의 현재 상태를 설명하는 텍스트를 생성합니다.

    Args:
        rows: 총 행 수
        cols: 총 열 수
        sample_data: 데이터 샘플 리스트
        formula_data: 수식 데이터 리스트

    Returns:
        엑셀 컨텍스트 설명 문자열
    """
    sample_text = "\n".join(sample_data) if sample_data else "데이터 없음"
    formula_text = "\n".join(formula_data) if formula_data else "수식 없음"

    return EXCEL_CONTEXT_TEMPLATE.format(
        rows=rows,
        cols=cols,
        sample_data=sample_text,
        formula_data=formula_text
    )