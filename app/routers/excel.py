from fastapi import APIRouter, UploadFile, File, HTTPException
from app.services.excel_service import process_excel_file

router = APIRouter()

@router.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    result = await process_excel_file(file)
    return {"message": "Excel processed successfully", "result": result}

@router.get("/download/{file_id}")
def download_excel(file_id: int):
    # TODO: 파일 다운로드 처리 구현
    return {"message": f"Download for file {file_id} not implemented yet."}