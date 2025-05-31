# app/services/excel.py
"""
엑셀 파일 조작 서비스
openpyxl을 사용하여 엑셀 파일을 직접 조작하는 기능을 제공합니다.
"""
import io
from typing import List, Any, Optional, Union
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter, column_index_from_string
import re

from app.schemas.excel import ExcelCommand


class ExcelManipulator:
    """
    엑셀 파일을 조작하는 클래스
    명령어를 받아서 실제 엑셀 파일을 수정합니다.
    """

    def __init__(self):
        """ExcelManipulator 초기화"""
        self.workbook: Optional[Workbook] = None
        self.active_sheet = None

    def load_from_bytes(self, excel_bytes: bytes) -> None:
        """
        바이트 데이터에서 엑셀 파일을 로드합니다.

        Args:
            excel_bytes: 엑셀 파일의 바이트 데이터
        """
        self.workbook = load_workbook(io.BytesIO(excel_bytes))
        self.active_sheet = self.workbook.active

    def save_to_bytes(self) -> bytes:
        """
        현재 워크북을 바이트 데이터로 저장합니다.

        Returns:
            엑셀 파일의 바이트 데이터
        """
        if not self.workbook:
            raise ValueError("워크북이 로드되지 않았습니다.")

        output = io.BytesIO()
        self.workbook.save(output)
        output.seek(0)
        return output.getvalue()

    def execute_commands(self, commands: List[ExcelCommand]) -> None:
        """
        명령어 리스트를 순차적으로 실행합니다.

        Args:
            commands: 실행할 ExcelCommand 리스트
        """
        if not self.workbook or not self.active_sheet:
            raise ValueError("워크북이 로드되지 않았습니다.")

        for command in commands:
            self._execute_single_command(command)

    def _execute_single_command(self, command: ExcelCommand) -> None:
        """
        단일 명령어를 실행합니다.

        Args:
            command: 실행할 ExcelCommand
        """
        command_type = command.command_type.lower()

        if command_type == "sum":
            self._apply_sum(command)
        elif command_type == "average":
            self._apply_average(command)
        elif command_type == "set_value":
            self._set_value(command)
        elif command_type == "bold":
            self._apply_bold(command)
        else:
            print(f"지원하지 않는 명령어: {command_type}")

    def _apply_sum(self, command: ExcelCommand) -> None:
        """
        SUM 함수를 적용합니다.

        Args:
            command: ExcelCommand (parameters에 범위가 포함되어야 함)
        """
        # parameters[0]에 범위가 있음 (예: "A1:A10")
        if command.parameters and "range" in command.parameters:
            range_str = command.parameters["range"]
            formula = f"=SUM({range_str})"
            self.active_sheet[command.target_range] = formula

    def _apply_average(self, command: ExcelCommand) -> None:
        """
        AVERAGE 함수를 적용합니다.

        Args:
            command: ExcelCommand (parameters에 범위가 포함되어야 함)
        """
        # parameters[0]에 범위가 있음 (예: "B1:B10")
        if command.parameters and "range" in command.parameters:
            range_str = command.parameters["range"]
            formula = f"=AVERAGE({range_str})"
            self.active_sheet[command.target_range] = formula

    def _set_value(self, command: ExcelCommand) -> None:
        """
        셀에 값을 설정합니다.

        Args:
            command: ExcelCommand (parameters에 설정할 값이 포함되어야 함)
        """
        if command.parameters and "value" in command.parameters:
            value = command.parameters["value"]

            # 범위인 경우 처리
            if ":" in command.target_range:
                # 범위의 모든 셀에 같은 값 설정
                for row in self.active_sheet[command.target_range]:
                    for cell in row:
                        cell.value = value
            else:
                # 단일 셀에 값 설정
                self.active_sheet[command.target_range] = value

    def _apply_bold(self, command: ExcelCommand) -> None:
        """
        셀에 굵은 글씨체를 적용합니다.

        Args:
            command: ExcelCommand (target_range에 적용할 범위가 포함되어야 함)
        """
        # 범위인 경우 처리
        if ":" in command.target_range:
            # 범위의 모든 셀에 bold 적용
            for row in self.active_sheet[command.target_range]:
                for cell in row:
                    cell.font = Font(bold=True)
        else:
            # 단일 셀에 bold 적용
            cell = self.active_sheet[command.target_range]
            cell.font = Font(bold=True)

    def _parse_range(self, range_str: str) -> tuple:
        """
        셀 범위 문자열을 파싱합니다.

        Args:
            range_str: 셀 범위 (예: "A1:B10")

        Returns:
            (start_col, start_row, end_col, end_row) 튜플
        """
        # 정규표현식으로 셀 주소 파싱
        pattern = r'([A-Z]+)(\d+)'

        if ":" in range_str:
            start, end = range_str.split(":")
            start_match = re.match(pattern, start)
            end_match = re.match(pattern, end)

            if start_match and end_match:
                start_col = column_index_from_string(start_match.group(1))
                start_row = int(start_match.group(2))
                end_col = column_index_from_string(end_match.group(1))
                end_row = int(end_match.group(2))
                return (start_col, start_row, end_col, end_row)
        else:
            match = re.match(pattern, range_str)
            if match:
                col = column_index_from_string(match.group(1))
                row = int(match.group(2))
                return (col, row, col, row)

        raise ValueError(f"잘못된 셀 범위 형식: {range_str}")


def process_excel_with_commands(
        excel_bytes: bytes,
        commands: Any
) -> bytes:
    """
    엑셀 파일에 명령어를 적용하고 결과를 반환합니다.

    Args:
        excel_bytes: 원본 엑셀 파일의 바이트 데이터
        commands: 적용할 명령어 리스트

    Returns:
        수정된 엑셀 파일의 바이트 데이터
    """
    manipulator = ExcelManipulator()

    # 엑셀 파일 로드
    manipulator.load_from_bytes(excel_bytes)

    # 명령어 실행
    manipulator.execute_commands(commands)

    # 결과 저장 및 반환
    return manipulator.save_to_bytes()


def create_empty_excel() -> bytes:
    """
    빈 엑셀 파일을 생성합니다.

    Returns:
        빈 엑셀 파일의 바이트 데이터
    """
    workbook = Workbook()
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()