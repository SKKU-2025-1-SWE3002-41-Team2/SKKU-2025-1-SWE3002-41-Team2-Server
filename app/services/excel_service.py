import io
from typing import List, Optional
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

from app.schemas.excel_schemas import ExcelCommand
from app.services.excel_commands import CommandType, ExcelCommandMapping

class ExcelService:
    """엑셀 파일 처리 서비스"""

    def __init__(self):
        self.command_mapping = ExcelCommandMapping()

    def load_excel_from_bytes(self, excel_bytes: bytes) -> Workbook:
        """바이트 데이터에서 엑셀 워크북 로드"""

    def convert_json_to_excel_bytes(sheet_json: dict) -> bytes:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_json.get("sheet_name", "Sheet1")

        data = sheet_json.get("data", {})

        for row_key, row_values in data.items():
            row_index = int(row_key.replace("row_", ""))
            for col_index, value in enumerate(row_values, start=1):
                ws.cell(row=row_index, column=col_index, value=value)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    def save_excel_to_bytes(self, workbook: Workbook) -> bytes:
        """워크북을 바이트 데이터로 저장"""
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return output.getvalue()

    def execute_command(self, workbook: Workbook, command: ExcelCommand) -> bool:
        """단일 엑셀 명령어 실행"""
        try:
            ws = workbook.active  # 첫 번째 시트만 사용

            command_info = self.command_mapping.get_command_info(command.command_type)
            command_type = command_info.get('type')

            if command_type == 'formula':
                # 수식 적용 - 직접 셀 참조 사용
                self._apply_formula(ws, command)
            elif command_type == 'format':
                # 서식 적용
                self._apply_format(ws, command)
            elif command_type == 'data':
                # 데이터 적용
                self._apply_data(ws, command)

            return True

        except Exception as e:
            print(f"명령어 실행 중 오류 발생: {str(e)}")
            return False

    def _apply_formula(self, worksheet, command: ExcelCommand):
        """수식 적용 - 단순화된 버전"""
        # 수식에 사용할 범위
        formula_range = command.parameters.get('range', command.target_range) # 예: "A1:A10"
        formula_type = command.command_type.upper() # 예: "SUM", "AVERAGE" 등

        # 엑셀 수식 생성 및 직접 할당
        formula = f"={formula_type}({formula_range})" #formula 변수가 "=SUM(A1:A10)"와 같은 형태로 설정됨
        worksheet[command.target_range] = formula # 수식이 적용될 셀에 직접 할당

    def _apply_format(self, worksheet, command: ExcelCommand):
        """서식 적용 - 범위에 대해 처리"""
        # 단일 셀 또는 범위 처리
        # command.command_type는 CommandType 클래스의 값 중 하나로, 예: "BOLD", "ITALIC" 등
        if ':' in command.target_range:
            # 범위인 경우
            for row in worksheet[command.target_range]:
                for cell in row:
                    self._format_cell(cell, command)
        else:
            # 단일 셀인 경우
            cell = worksheet[command.target_range] #워크 시트에서 타겟이되는 셀 객체 가져오기
            self._format_cell(cell, command)

    def _format_cell(self, cell, command: ExcelCommand):
        """개별 셀에 서식 적용"""
        if command.command_type == CommandType.BOLD.value:
            font = cell.font.copy()
            cell.font = Font(
                bold=True,
                size=font.size,
                name=font.name,
                color=font.color
            )

        elif command.command_type == CommandType.ITALIC.value:
            font = cell.font.copy()
            cell.font = Font(
                italic=True,
                size=font.size,
                name=font.name,
                color=font.color
            )

        elif command.command_type == CommandType.UNDERLINE.value:
            font = cell.font.copy()
            cell.font = Font(
                underline='single',
                size=font.size,
                name=font.name,
                color=font.color
            )

        elif command.command_type == CommandType.FONT_COLOR.value:
            color = command.parameters.get('color', '000000')
            font = cell.font.copy()
            cell.font = Font(
                color=color,
                size=font.size,
                name=font.name,
                bold=font.bold,
                italic=font.italic
            )

        elif command.command_type == CommandType.FILL_COLOR.value:
            color = command.parameters.get('color', 'FFFFFF')
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

        elif command.command_type == CommandType.BORDER.value:
            style = command.parameters.get('style', 'thin')
            border = Border(
                left=Side(style=style),
                right=Side(style=style),
                top=Side(style=style),
                bottom=Side(style=style)
            )
            cell.border = border

        elif command.command_type == CommandType.FONT_SIZE.value:
            size = command.parameters.get('size', 11)
            font = cell.font.copy()
            cell.font = Font(
                size=size,
                name=font.name,
                color=font.color,
                bold=font.bold,
                italic=font.italic
            )

        elif command.command_type == CommandType.FONT_NAME.value:
            name = command.parameters.get('name', 'Arial')
            font = cell.font.copy()
            cell.font = Font(
                name=name,
                size=font.size,
                color=font.color,
                bold=font.bold,
                italic=font.italic
            )

        elif command.command_type in [CommandType.ALIGN_LEFT.value, CommandType.ALIGN_CENTER.value,
                                      CommandType.ALIGN_RIGHT.value]:
            horizontal = command.command_type.replace('align_', '')
            alignment = cell.alignment.copy()
            cell.alignment = Alignment(horizontal=horizontal, vertical=alignment.vertical)

        elif command.command_type in [CommandType.ALIGN_TOP.value, CommandType.ALIGN_MIDDLE.value,
                                      CommandType.ALIGN_BOTTOM.value]:
            vertical = command.command_type.replace('align_', '')
            alignment = cell.alignment.copy()
            cell.alignment = Alignment(horizontal=alignment.horizontal, vertical=vertical)

    def _apply_data(self, worksheet, command: ExcelCommand):
        """데이터 적용"""
        if command.command_type == CommandType.SET_VALUE.value:
            value = command.parameters.get('value', '')
            if ':' in command.target_range:
                # 범위인 경우
                for row in worksheet[command.target_range]:
                    for cell in row:
                        cell.value = value
            else:
                # 단일 셀인 경우
                worksheet[command.target_range] = value

        elif command.command_type == CommandType.CLEAR.value:
            if ':' in command.target_range:
                # 범위인 경우
                for row in worksheet[command.target_range]:
                    for cell in row:
                        cell.value = None
            else:
                # 단일 셀인 경우
                worksheet[command.target_range] = None

        elif command.command_type == CommandType.MERGE.value:
            worksheet.merge_cells(command.target_range)

        elif command.command_type == CommandType.UNMERGE.value:
            worksheet.unmerge_cells(command.target_range)

    def execute_command_sequence(self, excel_bytes: bytes, commands: List[ExcelCommand]) -> bytes:
        """명령어 시퀀스 실행"""
        workbook = self.load_excel_from_bytes(excel_bytes)

        for command in commands:
            success = self.execute_command(workbook, command)
            if not success:
                print(f"명령어 실행 실패: {command}")

        return self.save_excel_to_bytes(workbook)