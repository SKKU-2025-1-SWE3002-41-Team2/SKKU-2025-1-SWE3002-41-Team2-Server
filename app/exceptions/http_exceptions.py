# app/exceptions/http_exceptions.py
from fastapi import HTTPException, status


#### auth ####
UnauthorizedException = HTTPException(
    status_code=status.HTTP_401_UNAUTHORIZED,
    detail="Invalid username or password."
)

UserNotFoundException = HTTPException(
    status_code=status.HTTP_404_NOT_FOUND,
    detail="User not found."
)

#### Session ####
SessionNotFoundException = HTTPException(
    status_code=status.HTTP_404_NOT_FOUND,
    detail="No chat sessions found for this user."
)

EmptyMessageAndSheetException = HTTPException(
    status_code=status.HTTP_400_BAD_REQUEST,
    detail="Either message or sheetData must be provided."
)