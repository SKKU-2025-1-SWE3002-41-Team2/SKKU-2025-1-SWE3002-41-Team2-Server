from fastapi import APIRouter, Depends, status, Response
from sqlalchemy.orm import Session
from app.database import get_db_session
from app.services.auth import login
from app.schemas.auth import LoginRequest, LoginResponse

router = APIRouter()

@router.post(
    "/login",
    summary="login",
    status_code=status.HTTP_200_OK,
    response_model=LoginResponse,
    responses={
        200: {"description": "login success"},
        401: {"description": "login fail (Unauthorized)"},
    }
)
def login_route(data: LoginRequest, db: Session = Depends(get_db_session)):
    user = login(db, data.username, data.password)
    return LoginResponse(
        username=user.username,
        userId=user.id
    )