import Dict


class ExcelCommand (BaseModel):
    """
    엑셀 명령어를 나타내는 모델
    """
    command_type: str  # 명령어 타입 (예: 'sum', 'average', 'font_color' 등)
    target_range: str  # 적용할 셀 범위 (예: 'A1:B2')
    parameters: Dict[str, Any]  # 명령어에 필요한 파라미터 (예: {'color': '#FF0000'})