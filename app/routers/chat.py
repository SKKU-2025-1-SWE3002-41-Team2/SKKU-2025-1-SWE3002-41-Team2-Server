import base64

from fastapi import APIRouter, Depends, Query, status, Form, UploadFile, File
from sqlalchemy.orm import Session
from app.database import get_db_session
from app.exceptions.http_exceptions import EmptyMessageAndSheetException
from app.schemas.chat import *
from app.services.chat import get_sessions, create_session, \
    delete_session, modify_session, get_messages, save_message_and_response

router = APIRouter()

@router.get(
    "/sessions",
    response_model=list[ChatSessionResponse],
    responses={
        404: {"description": "No chat sessions found"},
    },
)
def get_sessions_route(userId: int = Query(...), db: Session = Depends(get_db_session)):
    return get_sessions(userId=userId, db=db)


@router.post(
    "/sessions/create",
    response_model=ChatSessionCreateResponse,
    status_code=status.HTTP_201_CREATED,
    summary="Create a session or add message/sheet",
    responses={
        201: {"description": "Session created or message/sheet data added successfully."},
        400: {"description": "Either message or sheetData must be provided."},
    }
)
async def create_session_route(
    userId: int = Form(...),
    message: Optional[str] = Form(None),
    sheetData: Optional[UploadFile] = File(None),
    db: Session = Depends(get_db_session)
):
    if message is None and sheetData is None:
        raise EmptyMessageAndSheetException

    # sheetData를 byte로 읽기
    file_bytes = await sheetData.read() if sheetData is not None else None

    return create_session(userId, message, file_bytes, db)

@router.post(
    "/sessions/{sessionId}/message",
    summary="Save user message and sheet data, get LLM response",
    response_model=LLMResponse,
    status_code=status.HTTP_200_OK,
    responses={
        200: {"description": "Message saved and LLM response returned"},
        404: {"description": "Chat session not found"},
        400: {"description": "Invalid message or sheet data"},
    }
)
async def send_message_route(
    sessionId: int,
    message: str = Form(...),
    sheetData: Optional[UploadFile] = File(None),
    db: Session = Depends(get_db_session)
):
    # sheetData를 bytes로 읽음
    file_bytes = await sheetData.read() if sheetData is not None else None

    return save_message_and_response(sessionId, message, file_bytes, db)

@router.delete(
    "/sessions/{sessionId}",
    summary="Delete a chat session and all related messages/sheet",
    status_code=status.HTTP_204_NO_CONTENT,
    responses={
        204: {"description": "Session deleted successfully."},
        404: {"description": "Chat session not found."},
    }
)
def delete_session_route(sessionId: int, db: Session = Depends(get_db_session)):
    delete_session(sessionId, db)

@router.put(
    "/sessions/{sessionId}",
    summary="Modify session name",
    response_model=ChatSessionResponse,
    responses={
        200: {"description": "Session updated successfully."},
        404: {"description": "Chat session not found."}
    }
)
def update_session_route(sessionId: int, req: ChatSessionUpdateRequest, db: Session = Depends(get_db_session)):
    return modify_session(sessionId, req.name, db)

@router.get(
    "/sessions/{sessionId}",
    response_model=ChatSessionWithMessagesResponse,
    summary="Get a chat session with all messages",
    responses={
        200: {"description": "Chat session with messages returned"},
        404: {"description": "Chat session not found"}
    }
)
def get_session_messages_route(sessionId: int, db: Session = Depends(get_db_session)):
    session = get_messages(sessionId, db)

    # sheet가 존재하면 base64 인코딩 수행
    encoded_sheet = (
        base64.b64encode(session.sheet.sheetData).decode('utf-8')
        if session.sheet else None
    )
    return ChatSessionWithMessagesResponse(
        sessionId=session.id,
        userId=session.userId,
        name=session.name,
        modifiedAt=session.modifiedAt,
        sheetData=encoded_sheet,
        messages=session.messages
    )

