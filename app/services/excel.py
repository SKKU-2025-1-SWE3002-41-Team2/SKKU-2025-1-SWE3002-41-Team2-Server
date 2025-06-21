# app/services/excel.py
"""
엑셀 파일 조작 서비스
openpyxl을 사용하여 엑셀 파일을 직접 조작하는 기능을 제공합니다.
"""
import io
from typing import List, Any, Optional
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import column_index_from_string
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

        # 함수 관련 명령어
        if command_type == "sum":
            self._apply_sum(command)
        elif command_type == "average":
            self._apply_average(command)
        elif command_type == "count":
            self._apply_count(command)
        elif command_type == "max":
            self._apply_max(command)
        elif command_type == "min":
            self._apply_min(command)

        # 서식 관련 명령어
        elif command_type == "bold":
            self._apply_bold(command)
        elif command_type == "italic":
            self._apply_italic(command)
        elif command_type == "underline":
            self._apply_underline(command)
        elif command_type == "font_color":
            self._apply_font_color(command)
        elif command_type == "fill_color":
            self._apply_fill_color(command)
        elif command_type == "border":
            self._apply_border(command)
        elif command_type == "font_size":
            self._apply_font_size(command)
        elif command_type == "font_name":
            self._apply_font_name(command)

        # 데이터 관련 명령어
        elif command_type == "set_value":
            self._set_value(command)
        elif command_type == "clear":
            self._clear_cells(command)
        elif command_type == "merge":
            self._merge_cells(command)
        elif command_type == "unmerge":
            self._unmerge_cells(command)

        # 정렬 관련 명령어
        elif command_type in ["align_left", "align_center", "align_right",
                              "align_top", "align_middle", "align_bottom"]:
            self._apply_alignment(command)
        elif command_type in ["concatenate", "&"]: self._apply_concatenate(command)
        elif command_type == "left":   self._apply_left(command)
        elif command_type == "right":  self._apply_right(command)
        elif command_type == "mid":    self._apply_mid(command)
        elif command_type == "len":    self._apply_len(command)
        elif command_type == "round":  self._apply_round(command)
        elif command_type == "isblank":self._apply_isblank(command)
        else:
            print(f"지원하지 않는 명령어: {command_type}")

    # 함수 관련 명령어 구현
    def _apply_sum(self, command: ExcelCommand) -> None:
        """SUM 함수를 적용합니다."""
        if command.parameters and "range" in command.parameters:
            range_str = command.parameters["range"]
            formula = f"=SUM({range_str})"
            self.active_sheet[command.target_range] = formula

    def _apply_average(self, command: ExcelCommand) -> None:
        """AVERAGE 함수를 적용합니다."""
        if command.parameters and "range" in command.parameters:
            range_str = command.parameters["range"]
            formula = f"=AVERAGE({range_str})"
            self.active_sheet[command.target_range] = formula

    def _apply_count(self, command: ExcelCommand) -> None:
        """COUNT 함수를 적용합니다."""
        if command.parameters and "range" in command.parameters:
            range_str = command.parameters["range"]
            formula = f"=COUNT({range_str})"
            self.active_sheet[command.target_range] = formula

    def _apply_max(self, command: ExcelCommand) -> None:
        """MAX 함수를 적용합니다."""
        if command.parameters and "range" in command.parameters:
            range_str = command.parameters["range"]
            formula = f"=MAX({range_str})"
            self.active_sheet[command.target_range] = formula

    def _apply_min(self, command: ExcelCommand) -> None:
        """MIN 함수를 적용합니다."""
        if command.parameters and "range" in command.parameters:
            range_str = command.parameters["range"]
            formula = f"=MIN({range_str})"
            self.active_sheet[command.target_range] = formula

    def _apply_concatenate(self, command: ExcelCommand):
        """CONCATENATE 함수를 적용합니다."""
        values = command.parameters.get("values", [])
        if not values:
            return
        # 각 값을 셀 참조나 문자열로 처리
        arg_str = ",".join(str(v) for v in values)
        self.active_sheet[command.target_range] = f"=CONCATENATE({arg_str})"

    def _apply_left(self, command: ExcelCommand):
        """LEFT 함수를 적용합니다."""
        text = command.parameters.get("text", "")
        num_chars = command.parameters.get("num_chars", 1)
        if not text:
            return
        self.active_sheet[command.target_range] = f"=LEFT({text},{num_chars})"

    def _apply_right(self, command: ExcelCommand):
        """RIGHT 함수를 적용합니다."""
        text = command.parameters.get("text", "")
        num_chars = command.parameters.get("num_chars", 1)
        if not text:
            return
        self.active_sheet[command.target_range] = f"=RIGHT({text},{num_chars})"

    def _apply_mid(self, command: ExcelCommand):
        """MID 함수를 적용합니다."""
        text = command.parameters.get("text", "")
        start_num = command.parameters.get("start_num", 1)
        num_chars = command.parameters.get("num_chars", 1)
        if not text:
            return
        self.active_sheet[command.target_range] = f"=MID({text},{start_num},{num_chars})"

    def _apply_len(self, command: ExcelCommand):
        """LEN 함수를 적용합니다."""
        text = command.parameters.get("text", "")
        if not text:
            return
        self.active_sheet[command.target_range] = f"=LEN({text})"

    def _apply_round(self, command: ExcelCommand):
        """ROUND 함수를 적용합니다."""
        number = command.parameters.get("number", "")
        num_digits = command.parameters.get("num_digits", 0)
        if not number:
            return
        self.active_sheet[command.target_range] = f"=ROUND({number},{num_digits})"

    def _apply_isblank(self, command: ExcelCommand):
        """ISBLANK 함수를 적용합니다."""
        value = command.parameters.get("value", "")
        if not value:
            return
        self.active_sheet[command.target_range] = f"=ISBLANK({value})"

    # 서식 관련 명령어 구현
    def _apply_bold(self, command: ExcelCommand) -> None:
        """굵은 글씨체를 적용합니다."""
        self._apply_font_style(command.target_range, bold=True)

    def _apply_italic(self, command: ExcelCommand) -> None:
        """기울임체를 적용합니다."""
        self._apply_font_style(command.target_range, italic=True)

    def _apply_underline(self, command: ExcelCommand) -> None:
        """밑줄을 적용합니다."""
        self._apply_font_style(command.target_range, underline='single')

    def _apply_font_color(self, command: ExcelCommand) -> None:
        """글자 색상을 적용합니다."""
        if command.parameters and "color" in command.parameters:
            color = command.parameters["color"]
            self._apply_font_style(command.target_range, color=color)

    def _apply_fill_color(self, command: ExcelCommand) -> None:
        """배경색을 적용합니다."""
        if command.parameters and "color" in command.parameters:
            color = command.parameters["color"]
            fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
            self._apply_to_range(command.target_range, lambda cell: setattr(cell, 'fill', fill))

    def _apply_border(self, command: ExcelCommand) -> None:
        """테두리를 적용합니다."""
        style = "thin"  # 기본값
        if command.parameters and "style" in command.parameters:
            style = command.parameters["style"]

        border = Border(
            left=Side(style=style),
            right=Side(style=style),
            top=Side(style=style),
            bottom=Side(style=style)
        )
        self._apply_to_range(command.target_range, lambda cell: setattr(cell, 'border', border))

    def _apply_font_size(self, command: ExcelCommand) -> None:
        """글자 크기를 적용합니다."""
        if command.parameters and "size" in command.parameters:
            size = int(command.parameters["size"])
            self._apply_font_style(command.target_range, size=size)

    def _apply_font_name(self, command: ExcelCommand) -> None:
        """글꼴을 적용합니다."""
        if command.parameters and "name" in command.parameters:
            font_name = command.parameters["name"]
            self._apply_font_style(command.target_range, name=font_name)

    # 데이터 관련 명령어 구현
    def _set_value(self, command: ExcelCommand) -> None:
        """셀에 값을 설정합니다."""
        if command.parameters and "value" in command.parameters:
            value = command.parameters["value"]

            if ":" in command.target_range:
                # 범위의 모든 셀에 같은 값 설정
                self._apply_to_range(command.target_range, lambda cell: setattr(cell, 'value', value))
            else:
                # 단일 셀에 값 설정
                self.active_sheet[command.target_range] = value

    def _clear_cells(self, command: ExcelCommand) -> None:
        """셀의 내용을 지웁니다."""
        self._apply_to_range(command.target_range, lambda cell: setattr(cell, 'value', None))

    def _merge_cells(self, command: ExcelCommand) -> None:
        """셀을 병합합니다."""
        self.active_sheet.merge_cells(command.target_range)

    def _unmerge_cells(self, command: ExcelCommand) -> None:
        """셀 병합을 해제합니다."""
        self.active_sheet.unmerge_cells(command.target_range)

    # 정렬 관련 명령어 구현
    def _apply_alignment(self, command: ExcelCommand) -> None:
        """정렬을 적용합니다."""
        command_type = command.command_type.lower()

        # 현재 정렬 설정을 가져옴
        def apply_align(cell):
            current = cell.alignment.copy() if cell.alignment else Alignment()

            if command_type == "align_left":
                cell.alignment = Alignment(horizontal='left', vertical=current.vertical)
            elif command_type == "align_center":
                cell.alignment = Alignment(horizontal='center', vertical=current.vertical)
            elif command_type == "align_right":
                cell.alignment = Alignment(horizontal='right', vertical=current.vertical)
            elif command_type == "align_top":
                cell.alignment = Alignment(horizontal=current.horizontal, vertical='top')
            elif command_type == "align_middle":
                cell.alignment = Alignment(horizontal=current.horizontal, vertical='center')
            elif command_type == "align_bottom":
                cell.alignment = Alignment(horizontal=current.horizontal, vertical='bottom')

        self._apply_to_range(command.target_range, apply_align)

    # 헬퍼 메서드
    def _apply_font_style(self, target_range: str, **kwargs) -> None:
        """폰트 스타일을 적용하는 헬퍼 메서드"""

        def apply_font(cell):
            current_font = cell.font.__copy__() if cell.font else Font()

            # 현재 폰트의 속성을 유지하면서 새로운 속성만 업데이트
            font_dict = {
                'name': current_font.name,
                'size': current_font.size,
                'bold': current_font.bold,
                'italic': current_font.italic,
                'underline': current_font.underline,
                'color': current_font.color
            }

            # kwargs로 전달된 속성만 업데이트
            font_dict.update(kwargs)

            cell.font = Font(**font_dict)

        self._apply_to_range(target_range, apply_font)

    def _apply_to_range(self, target_range: str, func) -> None:
        """범위의 모든 셀에 함수를 적용하는 헬퍼 메서드"""
        if ":" in target_range:
            # 범위인 경우
            for row in self.active_sheet[target_range]:
                for cell in row:
                    func(cell)
        else:
            # 단일 셀인 경우
            cell = self.active_sheet[target_range]
            func(cell)

    def _parse_range(self, range_str: str) -> tuple:
        """
        셀 범위 문자열을 파싱합니다.

        Args:
            range_str: 셀 범위 (예: "A1:B10")

        Returns:
            (start_col, start_row, end_col, end_row) 튜플
        """
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