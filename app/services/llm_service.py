import os
from dotenv import load_dotenv
from openai import OpenAI
from pydantic import BaseModel
from typing import List, Literal, Optional
import json

'''
구현 이전에 추상적 계획 단계

4.1 프롬포팅 보다는 이전 공식 문서의 core concepts의 structured outputs 파트를 참조해서 작성함

chain of thought 예시에서와 같이
python class를 이용해 요청(아직 시도해보진 못했습니다.)

'''

# Pydantic 모델 정의 (응답 스키마)
# 클로드로 생성한 거 일부 편집
class ExcelFunction(BaseModel):
    """단일 엑셀 함수를 나타내는 모델"""
    #funtion_type : 여기서 Literal이 excel_service 단으로 넘어가서 실제 엑셀 함수 실행 시퀀스가 되기 위한 구분값으로 사용
    function_type: Literal["SUM", "AVERAGE", "VLOOKUP", "COUNTIF", "MAX", "MIN", "IF", "CONCATENATE"]
    target_cell: str  # 결과가 들어갈 셀 (예: "C3")
    range: Optional[str] = None  # 범위 (예: "A1:A10")
    """아래와 같이 여러 변수가 필요한 매개변수 별로 조직해서 추가하고, 구성해야됩니다.
    아마 이쪽 작업에서 가장 시간이 소요될 가능성이 큰 작업으로 보입니다. 
    아예 따로 class를 만들어 추가하는 형태로 구현해서 관리에 더 용이하도록 하는 것이 좋아보입니다.
    관련되서 claude가 만들어준 코드를 ~/llm_exam.py와 ~/llm_exam_service.py로 추가해두었습니다."""
    lookup_value: Optional[str] = None  # VLOOKUP용
    table_array: Optional[str] = None   # VLOOKUP용
    col_index_num: Optional[int] = None # VLOOKUP용
    range_lookup: Optional[bool] = None # VLOOKUP용
    criteria: Optional[str] = None      # COUNTIF용
    condition: Optional[str] = None     # IF용
    value_if_true: Optional[str] = None # IF용
    value_if_false: Optional[str] = None # IF용
    text_values: Optional[List[str]] = None # CONCATENATE용

class ExcelFunctionSequence(BaseModel):
    """엑셀 함수들의 시퀀스를 나타내는 모델"""
    functions: List[ExcelFunction]
    explanation: str  # 수행될 작업에 대한 설명

def get_excel_functions_structured(user_input: str, excel_context: str = None) -> ExcelFunctionSequence:
    """
    Structured Output을 사용하여 자연어를 엑셀 함수 시퀀스로 변환

    Args:
        user_input: 사용자의 자연어 명령
        excel_context: 현재 엑셀 시트 정보

    Returns:
        정형화된 엑셀 함수 시퀀스
    """

    system_prompt = """
    당신은 엑셀 전문가입니다. 사용자의 자연어 명령을 분석하여 적절한 엑셀 함수 시퀀스를 생성해주세요.

    지원하는 함수들:
    - SUM: 합계 계산
    - AVERAGE: 평균 계산  
    - VLOOKUP: 테이블에서 값 검색
    - COUNTIF: 조건에 맞는 셀 개수 계산
    - MAX: 최댓값 찾기
    - MIN: 최솟값 찾기
    - IF: 조건부 값 반환
    - CONCATENATE: 텍스트 연결

    주의사항:
    1. target_cell은 반드시 정확한 셀 주소로 지정 (예: "A1", "C10")
    2. range는 엑셀 형식으로 지정 (예: "A1:A10", "B:B", "A1:C20")
    3. 여러 단계 작업이 필요하면 순서대로 functions 배열에 포함
    4. 각 함수는 이전 함수의 결과를 참조할 수 있음 -
    """
    #4번 부분이 중요하게 작용될지 의논이 필요합니다

    user_message = f"사용자 요청: {user_input}"
    if excel_context:
        user_message += f"\n\n현재 엑셀 상황: {excel_context}"

    try:
        response = client.beta.chat.completions.parse(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message}
            ],
            response_format=ExcelFunctionSequence,
            temperature=0.1
        )

        return response.choices[0].message.parsed

    except Exception as e:
        print(f"Structured Output 호출 중 오류: {e}")
        # 기본값 반환
        return ExcelFunctionSequence(
            functions=[],
            explanation="오류가 발생하여 함수를 생성할 수 없습니다."
        )


def excel_function_to_formula(func: ExcelFunction) -> str:
    """ExcelFunction 객체를 실제 엑셀 수식으로 변환"""

    if func.function_type == "SUM":
        return f"=SUM({func.range})"

    elif func.function_type == "AVERAGE":
        return f"=AVERAGE({func.range})"

    elif func.function_type == "VLOOKUP":
        return f"=VLOOKUP({func.lookup_value},{func.table_array},{func.col_index_num},{func.range_lookup})"

    elif func.function_type == "COUNTIF":
        return f"=COUNTIF({func.range},\"{func.criteria}\")"

    elif func.function_type == "MAX":
        return f"=MAX({func.range})"

    elif func.function_type == "MIN":
        return f"=MIN({func.range})"

    elif func.function_type == "IF":
        return f"=IF({func.condition},{func.value_if_true},{func.value_if_false})"

    elif func.function_type == "CONCATENATE":
        values = ",".join(func.text_values) if func.text_values else ""
        return f"=CONCATENATE({values})"

    else:
        return f"알 수 없는 함수: {func.function_type}"


def process_excel_commands(user_input: str, excel_context: str = None):
    """
    사용자 명령을 처리하고 실행 가능한 형태로 변환하는 메인 함수
    """
    print(f"사용자 입력: {user_input}")
    print("=" * 50)

    # Structured Output으로 함수 시퀀스 생성
    function_sequence = get_excel_functions_structured(user_input, excel_context)

    print(f"설명: {function_sequence.explanation}")
    print(f"생성된 함수 개수: {len(function_sequence.functions)}")
    print()

    # 각 함수를 순서대로 처리
    for i, func in enumerate(function_sequence.functions, 1):
        print(f"단계 {i}:")
        print(f"  - 함수 타입: {func.function_type}")
        print(f"  - 대상 셀: {func.target_cell}")
        print(f"  - 엑셀 수식: {excel_function_to_formula(func)}")

        # 백엔드에서 실제 엑셀 조작을 위한 정보
        function_info = {
            "step": i,
            "target_cell": func.target_cell,
            "excel_formula": excel_function_to_formula(func),
            "function_data": func.dict()  # 원본 데이터도 포함
        }

        print(f"  - 백엔드 전송 데이터: {json.dumps(function_info, indent=2, ensure_ascii=False)}")
        print()

    return function_sequence



# 루트 디렉토리의 .env 파일 로드
load_dotenv()

async def process_natural_language_command(command: str) -> str:
    api_key = os.getenv("LLM_API_KEY")
    if not api_key:
        raise ValueError("LLM_API_KEY is not set in environment variables.")

    client = OpenAI(api_key=api_key)

    response = await client.responses.create(
        #추후 구현
        #공식 문서 참조 필요함
        #https://platform.openai.com/docs/guides/text?api-mode=responses
    )

    return response
