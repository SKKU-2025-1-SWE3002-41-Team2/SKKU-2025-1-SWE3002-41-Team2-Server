from fastapi import APIRouter
from . import chat, excel, llm, auth

router = APIRouter()

router.include_router(auth.router, prefix="/auth")
router.include_router(chat.router, prefix="/chat")
router.include_router(excel.router, prefix="/excel")