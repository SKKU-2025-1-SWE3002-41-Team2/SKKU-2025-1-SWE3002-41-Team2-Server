from enum import Enum
from typing import Dict, Any

class CommandType(Enum):
    """엑셀 명령어 타입"""
    # 함수 관련
    SUM = "sum"
    AVERAGE = "average"
    COUNT = "count"
    MAX = "max"
    MIN = "min"

    # 서식 관련
    BOLD = "bold"
    ITALIC = "italic"
    UNDERLINE = "underline"
    FONT_COLOR = "font_color"
    FILL_COLOR = "fill_color"
    BORDER = "border"
    FONT_SIZE = "font_size"
    FONT_NAME = "font_name"

    # 데이터 관련
    SET_VALUE = "set_value"
    CLEAR = "clear"
    MERGE = "merge"
    UNMERGE = "unmerge"

    # 정렬 관련
    ALIGN_LEFT = "align_left"
    ALIGN_CENTER = "align_center"
    ALIGN_RIGHT = "align_right"
    ALIGN_TOP = "align_top"
    ALIGN_MIDDLE = "align_middle"
    ALIGN_BOTTOM = "align_bottom"


class ExcelCommandMapping:
    """엑셀 명령어 매핑"""

    @staticmethod
    def get_command_info(command_type: str) -> Dict[str, Any]:
        """명령어 타입에 따른 정보 반환"""
        mapping = {
            CommandType.SUM.value: {
                "function": "SUM",
                "type": "formula",
                "description": "범위의 합계를 계산"
            },
            CommandType.AVERAGE.value: {
                "function": "AVERAGE",
                "type": "formula",
                "description": "범위의 평균을 계산"
            },
            CommandType.COUNT.value: {
                "function": "COUNT",
                "type": "formula",
                "description": "범위의 개수를 계산"
            },
            CommandType.MAX.value: {
                "function": "MAX",
                "type": "formula",
                "description": "범위의 최대값을 계산"
            },
            CommandType.MIN.value: {
                "function": "MIN",
                "type": "formula",
                "description": "범위의 최소값을 계산"
            },
            CommandType.BOLD.value: {
                "property": "bold",
                "type": "format",
                "description": "굵은 글씨체 적용"
            },
            CommandType.FONT_COLOR.value: {
                "property": "font.color",
                "type": "format",
                "description": "글자 색상 변경"
            },
            CommandType.FILL_COLOR.value: {
                "property": "fill",
                "type": "format",
                "description": "셀 배경색 변경"
            },
            CommandType.BORDER.value: {
                "property": "border",
                "type": "format",
                "description": "셀 테두리 설정"
            },
            CommandType.SET_VALUE.value: {
                "type": "data",
                "description": "셀 값 설정"
            }
        }
        return mapping.get(command_type, {})