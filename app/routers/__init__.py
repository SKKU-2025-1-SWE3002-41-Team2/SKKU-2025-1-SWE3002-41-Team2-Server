from fastapi import APIRouter
from . import chat, excel, llm, auth

router = APIRouter()

router.include_router(auth.router, prefix="/auth")