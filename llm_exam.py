from pydantic import BaseModel, Field
from typing import List, Literal, Optional, Union
from abc import ABC, abstractmethod


# =============================================================================
# 기본 매개변수 클래스들
# =============================================================================

class BaseFunctionParameters(BaseModel, ABC):
    """모든 엑셀 함수 매개변수의 기본 클래스"""

    @abstractmethod
    def to_excel_formula(self, target_cell: str) -> str:
        """엑셀 수식 문자열로 변환하는 추상 메서드"""
        pass

    @abstractmethod
    def validate_parameters(self) -> bool:
        """매개변수 유효성 검사 추상 메서드"""
        pass


# =============================================================================
# 범위 기반 함수들 (SUM, AVERAGE, MAX, MIN)
# =============================================================================

class RangeFunctionParameters(BaseFunctionParameters):
    """범위 기반 함수들의 공통 매개변수"""
    range: str = Field(..., description="계산할 셀 범위 (예: A1:A10, B:B)")

    def validate_parameters(self) -> bool:
        """범위가 유효한지 검사"""
        if not self.range or not self.range.strip():
            return False
        # 간단한 범위 형식 검증 (A1:B10, A:A 등)
        import re
        pattern = r'^[A-Z]+\d*:[A-Z]+\d*$|^[A-Z]+:[A-Z]+$'
        return bool(re.match(pattern, self.range.strip()))


class SumParameters(RangeFunctionParameters):
    """SUM 함수 매개변수"""

    def to_excel_formula(self, target_cell: str) -> str:
        return f"=SUM({self.range})"


class AverageParameters(RangeFunctionParameters):
    """AVERAGE 함수 매개변수"""

    def to_excel_formula(self, target_cell: str) -> str:
        return f"=AVERAGE({self.range})"


class MaxParameters(RangeFunctionParameters):
    """MAX 함수 매개변수"""

    def to_excel_formula(self, target_cell: str) -> str:
        return f"=MAX({self.range})"


class MinParameters(RangeFunctionParameters):
    """MIN 함수 매개변수"""

    def to_excel_formula(self, target_cell: str) -> str:
        return f"=MIN({self.range})"


# =============================================================================
# VLOOKUP 함수
# =============================================================================

class VlookupParameters(BaseFunctionParameters):
    """VLOOKUP 함수 매개변수"""
    lookup_value: str = Field(..., description="찾을 값이 있는 셀 또는 직접 값 (예: A1, 'John')")
    table_array: str = Field(..., description="검색할 테이블 범위 (예: A1:D100)")
    col_index_num: int = Field(..., ge=1, description="반환할 열의 인덱스 번호 (1부터 시작)")
    range_lookup: bool = Field(False, description="정확히 일치하는 값 찾기(False) 또는 근사치 허용(True)")

    def to_excel_formula(self, target_cell: str) -> str:
        # lookup_value가 셀 참조가 아닌 문자열 값인 경우 따옴표로 감싸기
        lookup_val = self.lookup_value
        if not lookup_val.startswith(('A', 'B', 'C', 'D', 'E', 'F', 'G')) and not lookup_val.startswith('='):
            lookup_val = f'"{lookup_val}"'

        return f"=VLOOKUP({lookup_val},{self.table_array},{self.col_index_num},{str(self.range_lookup).upper()})"

    def validate_parameters(self) -> bool:
        """VLOOKUP 매개변수 유효성 검사"""
        if not self.lookup_value or not self.table_array:
            return False
        if self.col_index_num < 1:
            return False

        # 테이블 범위 형식 검증
        import re
        pattern = r'^[A-Z]+\d+:[A-Z]+\d+$'
        return bool(re.match(pattern, self.table_array.strip()))


# =============================================================================
# COUNTIF 함수
# =============================================================================

class CountifParameters(BaseFunctionParameters):
    """COUNTIF 함수 매개변수"""
    range: str = Field(..., description="조건을 확인할 셀 범위 (예: A1:A10)")
    criteria: str = Field(..., description="조건 (예: '>50', 'John', '>=100')")

    def to_excel_formula(self, target_cell: str) -> str:
        # 조건이 숫자 비교가 아닌 문자열인 경우 따옴표 처리
        criteria_formatted = self.criteria
        if not any(op in self.criteria for op in ['>', '<', '=', '>=', '<=', '<>']):
            # 단순 문자열인 경우
            if not self.criteria.startswith('"') and not self.criteria.endswith('"'):
                criteria_formatted = f'"{self.criteria}"'

        return f"=COUNTIF({self.range},{criteria_formatted})"

    def validate_parameters(self) -> bool:
        """COUNTIF 매개변수 유효성 검사"""
        if not self.range or not self.criteria:
            return False

        # 범위 형식 검증
        import re
        pattern = r'^[A-Z]+\d*:[A-Z]+\d*$|^[A-Z]+:[A-Z]+$'
        return bool(re.match(pattern, self.range.strip()))


# =============================================================================
# IF 함수
# =============================================================================

class IfParameters(BaseFunctionParameters):
    """IF 함수 매개변수"""
    condition: str = Field(..., description="조건식 (예: A1>50, B2='Pass')")
    value_if_true: str = Field(..., description="조건이 참일 때 반환할 값")
    value_if_false: str = Field(..., description="조건이 거짓일 때 반환할 값")

    def to_excel_formula(self, target_cell: str) -> str:
        # 문자열 값들에 대한 따옴표 처리
        true_val = self._format_value(self.value_if_true)
        false_val = self._format_value(self.value_if_false)

        return f"=IF({self.condition},{true_val},{false_val})"

    def _format_value(self, value: str) -> str:
        """값의 타입에 따라 적절한 형식으로 변환"""
        # 셀 참조인 경우 그대로 반환
        if value.startswith(('A', 'B', 'C', 'D', 'E', 'F', 'G')) or value.startswith('='):
            return value

        # 숫자인지 확인
        try:
            float(value)
            return value
        except ValueError:
            # 문자열인 경우 따옴표로 감싸기
            if not value.startswith('"') and not value.endswith('"'):
                return f'"{value}"'
            return value

    def validate_parameters(self) -> bool:
        """IF 매개변수 유효성 검사"""
        return all([self.condition, self.value_if_true, self.value_if_false])


# =============================================================================
# CONCATENATE 함수
# =============================================================================

class ConcatenateParameters(BaseFunctionParameters):
    """CONCATENATE 함수 매개변수"""
    text_values: List[str] = Field(..., min_items=1, description="연결할 텍스트 값들의 리스트")

    def to_excel_formula(self, target_cell: str) -> str:
        # 각 값을 적절한 형식으로 변환
        formatted_values = []
        for value in self.text_values:
            # 셀 참조인 경우 그대로 사용
            if value.startswith(('A', 'B', 'C', 'D', 'E', 'F', 'G')) or value.startswith('='):
                formatted_values.append(value)
            else:
                # 문자열인 경우 따옴표로 감싸기
                if not value.startswith('"') and not value.endswith('"'):
                    formatted_values.append(f'"{value}"')
                else:
                    formatted_values.append(value)

        return f"=CONCATENATE({','.join(formatted_values)})"

    def validate_parameters(self) -> bool:
        """CONCATENATE 매개변수 유효성 검사"""
        return len(self.text_values) > 0 and all(val.strip() for val in self.text_values)


# =============================================================================
# 통합 ExcelFunction 클래스
# =============================================================================

# 모든 매개변수 타입의 Union
ExcelFunctionParams = Union[
    SumParameters,
    AverageParameters,
    MaxParameters,
    MinParameters,
    VlookupParameters,
    CountifParameters,
    IfParameters,
    ConcatenateParameters
]


class ExcelFunction(BaseModel):
    """단일 엑셀 함수를 나타내는 모델"""
    function_type: Literal["SUM", "AVERAGE", "VLOOKUP", "COUNTIF", "MAX", "MIN", "IF", "CONCATENATE"]
    target_cell: str = Field(..., description="결과가 들어갈 셀 (예: C3)")
    parameters: ExcelFunctionParams = Field(..., description="함수별 매개변수")

    def get_excel_formula(self) -> str:
        """엑셀 수식 문자열 반환"""
        return self.parameters.to_excel_formula(self.target_cell)

    def validate_function(self) -> bool:
        """함수와 매개변수의 유효성 검사"""
        # 함수 타입과 매개변수 타입이 일치하는지 확인
        type_mapping = {
            "SUM": SumParameters,
            "AVERAGE": AverageParameters,
            "MAX": MaxParameters,
            "MIN": MinParameters,
            "VLOOKUP": VlookupParameters,
            "COUNTIF": CountifParameters,
            "IF": IfParameters,
            "CONCATENATE": ConcatenateParameters
        }

        expected_type = type_mapping.get(self.function_type)
        if not isinstance(self.parameters, expected_type):
            return False

        return self.parameters.validate_parameters()


class ExcelFunctionSequence(BaseModel):
    """엑셀 함수들의 시퀀스를 나타내는 모델"""
    functions: List[ExcelFunction]
    explanation: str = Field(..., description="수행될 작업에 대한 설명")

    def validate_sequence(self) -> bool:
        """전체 시퀀스의 유효성 검사"""
        if not self.functions:
            return False

        return all(func.validate_function() for func in self.functions)

    def get_execution_plan(self) -> List[dict]:
        """실행 계획을 딕셔너리 리스트로 반환"""
        execution_plan = []
        for i, func in enumerate(self.functions, 1):
            execution_plan.append({
                "step": i,
                "function_type": func.function_type,
                "target_cell": func.target_cell,
                "excel_formula": func.get_excel_formula(),
                "parameters": func.parameters.dict(),
                "is_valid": func.validate_function()
            })
        return execution_plan


# =============================================================================
# 팩토리 함수들 (편의성을 위한)
# =============================================================================

class ExcelFunctionFactory:
    """엑셀 함수 객체를 쉽게 생성하기 위한 팩토리 클래스"""

    @staticmethod
    def create_sum(target_cell: str, range_str: str) -> ExcelFunction:
        """SUM 함수 생성"""
        return ExcelFunction(
            function_type="SUM",
            target_cell=target_cell,
            parameters=SumParameters(range=range_str)
        )

    @staticmethod
    def create_average(target_cell: str, range_str: str) -> ExcelFunction:
        """AVERAGE 함수 생성"""
        return ExcelFunction(
            function_type="AVERAGE",
            target_cell=target_cell,
            parameters=AverageParameters(range=range_str)
        )

    @staticmethod
    def create_vlookup(target_cell: str, lookup_value: str, table_array: str,
                       col_index_num: int, range_lookup: bool = False) -> ExcelFunction:
        """VLOOKUP 함수 생성"""
        return ExcelFunction(
            function_type="VLOOKUP",
            target_cell=target_cell,
            parameters=VlookupParameters(
                lookup_value=lookup_value,
                table_array=table_array,
                col_index_num=col_index_num,
                range_lookup=range_lookup
            )
        )

    @staticmethod
    def create_countif(target_cell: str, range_str: str, criteria: str) -> ExcelFunction:
        """COUNTIF 함수 생성"""
        return ExcelFunction(
            function_type="COUNTIF",
            target_cell=target_cell,
            parameters=CountifParameters(range=range_str, criteria=criteria)
        )

    @staticmethod
    def create_if(target_cell: str, condition: str, value_if_true: str,
                  value_if_false: str) -> ExcelFunction:
        """IF 함수 생성"""
        return ExcelFunction(
            function_type="IF",
            target_cell=target_cell,
            parameters=IfParameters(
                condition=condition,
                value_if_true=value_if_true,
                value_if_false=value_if_false
            )
        )

    @staticmethod
    def create_concatenate(target_cell: str, text_values: List[str]) -> ExcelFunction:
        """CONCATENATE 함수 생성"""
        return ExcelFunction(
            function_type="CONCATENATE",
            target_cell=target_cell,
            parameters=ConcatenateParameters(text_values=text_values)
        )


# =============================================================================
# 사용 예시
# =============================================================================

if __name__ == "__main__":
    # 팩토리를 사용한 함수 생성 예시
    factory = ExcelFunctionFactory()

    # 다양한 함수 생성
    sum_func = factory.create_sum("C1", "A1:A10")
    avg_func = factory.create_average("C2", "A1:A10")
    vlookup_func = factory.create_vlookup("D1", "B1", "A1:C100", 3, False)
    countif_func = factory.create_countif("E1", "A1:A10", ">50")
    if_func = factory.create_if("F1", "A1>50", "Pass", "Fail")
    concat_func = factory.create_concatenate("G1", ["Hello", " ", "World"])

    # 함수 시퀀스 생성
    sequence = ExcelFunctionSequence(
        functions=[sum_func, avg_func, vlookup_func, countif_func, if_func, concat_func],
        explanation="A열 데이터를 분석하고 다양한 계산을 수행합니다."
    )

    # 실행 계획 출력
    execution_plan = sequence.get_execution_plan()
    for step in execution_plan:
        print(f"단계 {step['step']}: {step['function_type']}")
        print(f"  대상 셀: {step['target_cell']}")
        print(f"  수식: {step['excel_formula']}")
        print(f"  유효성: {step['is_valid']}")
        print()

    # 유효성 검사
    print(f"전체 시퀀스 유효성: {sequence.validate_sequence()}")