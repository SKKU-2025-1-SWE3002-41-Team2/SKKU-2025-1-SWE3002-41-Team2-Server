import openai
from typing import List, Dict, Optional
import json
from pydantic import ValidationError

# 위에서 정의한 클래스들을 import (실제로는 별도 파일에서 import)
from excel_function_models import (
    ExcelFunction, ExcelFunctionSequence, ExcelFunctionFactory,
    SumParameters, AverageParameters, VlookupParameters, CountifParameters,
    IfParameters, ConcatenateParameters, MaxParameters, MinParameters
)


class ExcelLLMProcessor:
    """OpenAI API를 사용하여 자연어를 엑셀 함수로 변환하는 프로세서"""

    def __init__(self, api_key: str):
        self.client = openai.OpenAI(api_key=api_key)
        self.factory = ExcelFunctionFactory()

    def get_excel_functions_structured(self, user_command: str, excel_context: str = None) -> ExcelFunctionSequence:
        """
        Structured Output을 사용하여 자연어를 엑셀 함수 시퀀스로 변환
        """
        system_prompt = self._create_system_prompt()
        user_message = self._create_user_message(user_command, excel_context)

        try:
            response = self.client.beta.chat.completions.parse(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_message}
                ],
                response_format=ExcelFunctionSequence,
                temperature=0.1
            )

            function_sequence = response.choices[0].message.parsed

            # 생성된 함수들의 유효성 검사
            if not function_sequence.validate_sequence():
                print("⚠️  생성된 함수 시퀀스에 유효하지 않은 함수가 있습니다.")
                # 유효하지 않은 함수들을 로깅하거나 수정할 수 있음

            return function_sequence

        except Exception as e:
            print(f"Structured Output 호출 중 오류: {e}")
            # 기본값 반환
            return ExcelFunctionSequence(
                functions=[],
                explanation="오류가 발생하여 함수를 생성할 수 없습니다."
            )

    def get_excel_functions_function_calling(self, user_command: str,
                                             excel_context: str = None) -> ExcelFunctionSequence:
        """
        Function Calling을 사용하여 자연어를 엑셀 함수 시퀀스로 변환
        """
        functions = self._create_function_definitions()
        system_prompt = self._create_system_prompt_for_function_calling()
        user_message = self._create_user_message(user_command, excel_context)

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_message}
                ],
                functions=functions,
                function_call="auto",
                temperature=0.1
            )

            # Function calling 결과를 ExcelFunctionSequence로 변환
            return self._parse_function_calling_response(response)

        except Exception as e:
            print(f"Function Calling 호출 중 오류: {e}")
            return ExcelFunctionSequence(
                functions=[],
                explanation="오류가 발생하여 함수를 생성할 수 없습니다."
            )

    def _create_system_prompt(self) -> str:
        """Structured Output용 시스템 프롬프트"""
        return """
        당신은 엑셀 전문가입니다. 사용자의 자연어 명령을 분석하여 적절한 엑셀 함수 시퀀스를 생성해주세요.

        지원하는 함수들과 매개변수:

        1. SUM: 합계 계산
           - range: 합계할 셀 범위 (예: "A1:A10")

        2. AVERAGE: 평균 계산
           - range: 평균을 구할 셀 범위 (예: "A1:A10")

        3. MAX/MIN: 최댓값/최솟값
           - range: 값을 찾을 셀 범위 (예: "A1:A10")

        4. VLOOKUP: 테이블에서 값 검색
           - lookup_value: 찾을 값 (예: "A1" 또는 "John")
           - table_array: 검색할 테이블 범위 (예: "A1:D100")
           - col_index_num: 반환할 열 번호 (1부터 시작)
           - range_lookup: 정확히 일치 여부 (false 권장)

        5. COUNTIF: 조건에 맞는 셀 개수
           - range: 조건을 확인할 범위 (예: "A1:A10")
           - criteria: 조건 (예: ">50", "Pass")

        6. IF: 조건부 값 반환
           - condition: 조건식 (예: "A1>50")
           - value_if_true: 참일 때 값
           - value_if_false: 거짓일 때 값

        7. CONCATENATE: 텍스트 연결
           - text_values: 연결할 텍스트들의 배열

        주의사항:
        1. target_cell은 정확한 셀 주소로 지정 (예: "A1", "C10")
        2. range는 엑셀 형식으로 지정 (예: "A1:A10", "B:B")
        3. 여러 단계 작업이 필요하면 순서대로 functions 배열에 포함
        4. 각 함수의 매개변수는 해당 함수에 맞는 타입으로만 설정
        """

    def _create_system_prompt_for_function_calling(self) -> str:
        """Function Calling용 시스템 프롬프트"""
        return """
        당신은 엑셀 전문가입니다. 사용자의 자연어 명령을 분석하여 적절한 엑셀 함수를 호출해주세요.
        여러 개의 함수가 필요한 경우 순서대로 함수를 호출해주세요.

        주의사항:
        - 셀 범위는 정확한 엑셀 형식으로 지정해주세요 (예: A1:A10, B:B)
        - target_cell은 결과가 들어갈 정확한 셀 위치를 지정해주세요
        - 여러 단계의 작업이 필요하면 순서대로 함수를 호출해주세요
        """

    def _create_user_message(self, user_command: str, excel_context: str = None) -> str:
        """사용자 메시지 생성"""
        message = f"사용자 요청: {user_command}"
        if excel_context:
            message += f"\n\n현재 엑셀 상황: {excel_context}"
        return message

    def _create_function_definitions(self) -> List[Dict]:
        """Function Calling용 함수 정의들"""
        return [
            {
                "name": "create_sum_function",
                "description": "SUM 함수를 생성합니다",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "target_cell": {"type": "string", "description": "결과를 입력할 셀"},
                        "range": {"type": "string", "description": "합계할 셀 범위"}
                    },
                    "required": ["target_cell", "range"]
                }
            },
            {
                "name": "create_average_function",
                "description": "AVERAGE 함수를 생성합니다",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "target_cell": {"type": "string", "description": "결과를 입력할 셀"},
                        "range": {"type": "string", "description": "평균을 구할 셀 범위"}
                    },
                    "required": ["target_cell", "range"]
                }
            },
            {
                "name": "create_vlookup_function",
                "description": "VLOOKUP 함수를 생성합니다",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "target_cell": {"type": "string", "description": "결과를 입력할 셀"},
                        "lookup_value": {"type": "string", "description": "찾을 값"},
                        "table_array": {"type": "string", "description": "검색할 테이블 범위"},
                        "col_index_num": {"type": "integer", "description": "반환할 열 번호"},
                        "range_lookup": {"type": "boolean", "description": "정확히 일치 여부"}
                    },
                    "required": ["target_cell", "lookup_value", "table_array", "col_index_num"]
                }
            },
            {
                "name": "create_countif_function",
                "description": "COUNTIF 함수를 생성합니다",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "target_cell": {"type": "string", "description": "결과를 입력할 셀"},
                        "range": {"type": "string", "description": "조건을 확인할 범위"},
                        "criteria": {"type": "string", "description": "조건"}
                    },
                    "required": ["target_cell", "range", "criteria"]
                }
            },
            {
                "name": "create_if_function",
                "description": "IF 함수를 생성합니다",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "target_cell": {"type": "string", "description": "결과를 입력할 셀"},
                        "condition": {"type": "string", "description": "조건식"},
                        "value_if_true": {"type": "string", "description": "참일 때 값"},
                        "value_if_false": {"type": "string", "description": "거짓일 때 값"}
                    },
                    "required": ["target_cell", "condition", "value_if_true", "value_if_false"]
                }
            }
        ]

    def _parse_function_calling_response(self, response) -> ExcelFunctionSequence:
        """Function Calling 응답을 ExcelFunctionSequence로 변환"""
        functions = []
        explanation_parts = []

        message = response.choices[0].message

        if hasattr(message, 'function_call') and message.function_call:
            # 단일 함수 호출
            func = self._create_function_from_call(message.function_call)
            if func:
                functions.append(func)
                explanation_parts.append(f"{func.function_type} 함수를 {func.target_cell}에 생성")

        elif hasattr(message, 'tool_calls') and message.tool_calls:
            # 다중 함수 호출 (최신 API)
            for tool_call in message.tool_calls:
                if tool_call.type == 'function':
                    func = self._create_function_from_call(tool_call.function)
                    if func:
                        functions.append(func)
                        explanation_parts.append(f"{func.function_type} 함수를 {func.target_cell}에 생성")

        explanation = "; ".join(explanation_parts) if explanation_parts else "함수를 생성했습니다."

        return ExcelFunctionSequence(
            functions=functions,
            explanation=explanation
        )

    def _create_function_from_call(self, function_call) -> Optional[ExcelFunction]:
        """함수 호출 정보에서 ExcelFunction 객체 생성"""
        try:
            function_name = function_call.name
            arguments = json.loads(function_call.arguments)

            if function_name == "create_sum_function":
                return self.factory.create_sum(
                    arguments["target_cell"],
                    arguments["range"]
                )
            elif function_name == "create_average_function":
                return self.factory.create_average(
                    arguments["target_cell"],
                    arguments["range"]
                )
            elif function_name == "create_vlookup_function":
                return self.factory.create_vlookup(
                    arguments["target_cell"],
                    arguments["lookup_value"],
                    arguments["table_array"],
                    arguments["col_index_num"],
                    arguments.get("range_lookup", False)
                )
            elif function_name == "create_countif_function":
                return self.factory.create_countif(
                    arguments["target_cell"],
                    arguments["range"],
                    arguments["criteria"]
                )
            elif function_name == "create_if_function":
                return self.factory.create_if(
                    arguments["target_cell"],
                    arguments["condition"],
                    arguments["value_if_true"],
                    arguments["value_if_false"]
                )

        except (json.JSONDecodeError, KeyError, ValidationError) as e:
            print(f"함수 생성 중 오류: {e}")

        return None


# =============================================================================
# FastAPI 통합 예시
# =============================================================================

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import os

app = FastAPI(title="Excel LLM Platform - Updated")


class ExcelCommandRequest(BaseModel):
    """엑셀 명령 요청 모델"""
    chat_id: int
    user_command: str
    current_sheet_data: dict
    excel_context: Optional[str] = None
    use_structured_output: bool = True  # True: Structured Output, False: Function Calling


class ExcelCommandResponse(BaseModel):
    """엑셀 명령 응답 모델"""
    success: bool
    explanation: str
    functions: List[dict]  # 실행 계획
    updated_sheet_data: dict
    execution_results: List[dict]


# 전역 프로세서 인스턴스
processor = ExcelLLMProcessor(api_key=os.getenv("OPENAI_API_KEY"))


@app.post("/api/excel/process-command", response_model=ExcelCommandResponse)
async def process_excel_command(request: ExcelCommandRequest):
    """
    업데이트된 자연어 명령 처리 엔드포인트
    """
    try:
        # 1. 엑셀 컨텍스트 분석
        excel_context = request.excel_context or analyze_sheet_context(request.current_sheet_data)

        # 2. 자연어를 엑셀 함수로 변환 (방법 선택)
        if request.use_structured_output:
            function_sequence = processor.get_excel_functions_structured(
                request.user_command,
                excel_context
            )
        else:
            function_sequence = processor.get_excel_functions_function_calling(
                request.user_command,
                excel_context
            )

        if not function_sequence.functions:
            raise HTTPException(
                status_code=400,
                detail="유효한 엑셀 함수를 생성할 수 없습니다."
            )

        # 3. 실행 계획 생성
        execution_plan = function_sequence.get_execution_plan()

        # 4. 실제 엑셀 시트에 함수 적용 (여기서는 시뮬레이션)
        updated_sheet, execution_results = apply_functions_to_sheet(
            request.current_sheet_data,
            function_sequence.functions
        )

        return ExcelCommandResponse(
            success=True,
            explanation=function_sequence.explanation,
            functions=execution_plan,
            updated_sheet_data=updated_sheet,
            execution_results=execution_results
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"처리 중 오류: {str(e)}")


def analyze_sheet_context(sheet_data: dict) -> str:
    """시트 데이터를 분석하여 컨텍스트 생성 (구현 필요)"""
    # 실제 구현에서는 SheetJS 데이터를 분석
    return "현재 시트에 데이터가 있습니다."


def apply_functions_to_sheet(sheet_data: dict, functions: List[ExcelFunction]) -> tuple:
    """함수들을 시트에 적용 (실제 구현 필요)"""
    # 실제 구현에서는 Univer나 다른 라이브러리 사용
    updated_sheet = sheet_data.copy()
    execution_results = []

    for func in functions:
        execution_results.append({
            "function_type": func.function_type,
            "target_cell": func.target_cell,
            "formula": func.get_excel_formula(),
            "success": True,
            "message": f"{func.function_type} 함수가 {func.target_cell}에 적용되었습니다."
        })

    return updated_sheet, execution_results


# =============================================================================
# 테스트 코드
# =============================================================================

if __name__ == "__main__":
    # 테스트용
    import asyncio


    async def test_processor():
        test_processor = ExcelLLMProcessor("your-api-key-here")

        # Structured Output 테스트
        print("=== Structured Output 테스트 ===")
        result1 = test_processor.get_excel_functions_structured(
            "A열의 합계를 B1에, 평균을 B2에 넣어주세요",
            "A1:A10에 숫자 데이터가 있음"
        )

        print(f"설명: {result1.explanation}")
        execution_plan = result1.get_execution_plan()
        for step in execution_plan:
            print(f"  {step['step']}. {step['function_type']}: {step['excel_formula']}")

        print(f"유효성: {result1.validate_sequence()}")

        # Function Calling 테스트
        print("\n=== Function Calling 테스트 ===")
        result2 = test_processor.get_excel_functions_function_calling(
            "학생 점수에서 60점 이상인 학생 수를 C1에 넣어주세요",
            "B1:B20에 학생 점수가 있음"
        )

        print(f"설명: {result2.explanation}")
        execution_plan2 = result2.get_execution_plan()
        for step in execution_plan2:
            print(f"  {step['step']}. {step['function_type']}: {step['excel_formula']}")


    # asyncio.run(test_processor())
    print("프로세서가 설정되었습니다. FastAPI 서버를 실행하여 테스트하세요.")