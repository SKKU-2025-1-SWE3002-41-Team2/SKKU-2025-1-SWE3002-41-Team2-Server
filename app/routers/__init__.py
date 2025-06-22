from fastapi import APIRouter
from . import chat_router, auth_router, llm_router

router = APIRouter()

router.include_router(auth_router.router, prefix="/auth", tags=["auth"])
router.include_router(auth_router.router, prefix="/chat", tags=["chat"])
router.include_router(llm_router.router, prefix="/llm", tags=["llm"])