from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy.orm import Session
from app.database import get_db_session
from app.schemas.chat import *
from app.schemas.llm import LLMResponse
from app.services.chat import get_sessions, create_session, create_session, \
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
    response_model=LLMResponse,
    status_code=status.HTTP_201_CREATED,
    summary="Create a session or add message/sheet",
    responses={
        201: {"description": "Session created or message/sheet data added successfully."},
        400: {"description": "Either message or sheetData must be provided."},
    }
)
def create_session_route(data: ChatSessionCreateRequest, db: Session = Depends(get_db_session)):
    res = create_session(data, db)
    return res

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
    return ChatSessionWithMessagesResponse(
        sessionId=session.id,
        userId=session.userId,
        name=session.name,
        modifiedAt=session.modifiedAt,
        sheetData=session.sheet.sheetData if session.sheet else None,
        messages=session.messages
    )

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
def save_message_route(
    sessionId: int,
    data: MessageRequest,
    db: Session = Depends(get_db_session)
):
    return save_message_and_response(sessionId, data, db)