from fastapi import APIRouter, HTTPException

router = APIRouter()

@router.post("/command")
def process_command(command: str):
    # TODO: LLM과의 연동 처리 (예: OpenAI API 호출)
    response = {"response": f"Processed command: {command}"}
    return response