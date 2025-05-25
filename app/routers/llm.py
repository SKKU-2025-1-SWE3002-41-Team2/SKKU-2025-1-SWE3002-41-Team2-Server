# app/routers/llm.py
from fastapi import APIRouter, HTTPException
from app.services.llm_service import process_natural_language_command

router = APIRouter()

@router.post("/process-command/")
async def process_command(command: str):
    try:
        result = await process_natural_language_command(command)
        return {"result": result}
    except Exception as e:
        # LLM 연동 또는 기타 예외 처리
        raise HTTPException(status_code=500, detail=str(e))
