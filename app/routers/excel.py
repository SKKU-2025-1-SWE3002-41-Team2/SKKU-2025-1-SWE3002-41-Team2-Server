from fastapi import APIRouter, Response
from fastapi.responses import StreamingResponse
from io import BytesIO
import openpyxl

router = APIRouter()

@router.get("/download-excel/")
async def download_excel():
    # 샘플 Excel 파일 생성
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sample"
    ws.append(["Name", "Score"])
    ws.append(["Alice", 95])
    ws.append(["Bob", 88])

    # 파일을 메모리에 저장
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    # StreamingResponse로 반환
    headers = {
        'Content-Disposition': 'attachment; filename="sample.xlsx"'
    }
    return StreamingResponse(buffer, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers=headers)
