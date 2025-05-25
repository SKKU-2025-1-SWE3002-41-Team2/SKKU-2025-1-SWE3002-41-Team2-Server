# app/exceptions/http_exceptions.py
from fastapi import HTTPException, status

UserNotFoundException = HTTPException(
    status_code=status.HTTP_404_NOT_FOUND,
    detail="User not found."
)

SessionNotFoundException = HTTPException(
    status_code=status.HTTP_404_NOT_FOUND,
    detail="No chat sessions found for this user."
)

UnauthorizedException = HTTPException(
    status_code=status.HTTP_401_UNAUTHORIZED,
    detail="Invalid username or password."
)
