import pytest
from unittest.mock import MagicMock
from datetime import datetime
from app.services import chat
from app.models import ChatSession, Message, ChatSheet, User
from app.exceptions.http_exceptions import SessionNotFoundException, UserNotFoundException
from app.utils.timezone import KST

# 사용자의 세션 목록을 정상적으로 불러올 수 있는지 테스트
def test_get_sessions_success():
    mock_db = MagicMock()
    mock_db.query().filter().order_by().all.return_value = [
        ChatSession(id=1, userId=10, name="Test Session", summary="test", modifiedAt=datetime.now())
    ]
    sessions = chat.get_sessions(userId=10, db=mock_db)
    assert len(sessions) == 1
    assert sessions[0].userId == 10

# 세션이 없을 때 예외가 발생하는지 테스트
def test_get_sessions_not_found():
    mock_db = MagicMock()
    mock_db.query().filter().order_by().all.return_value = []
    with pytest.raises(SessionNotFoundException):
        chat.get_sessions(userId=99, db=mock_db)

# 특정 세션의 메시지를 정상 조회할 수 있는지 테스트
def test_get_messages_success():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = ChatSession(id=1, userId=1)
    session = chat.get_messages(session_id=1, db=mock_db)
    assert session.id == 1

# 존재하지 않는 세션 조회 시 예외 발생하는지 테스트
def test_get_messages_not_found():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None
    with pytest.raises(SessionNotFoundException):
        chat.get_messages(session_id=1, db=mock_db)

# 유저가 존재할 때 정상적으로 통과되는지 테스트
def test_validate_user_exists_success():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = User(id=1)
    assert chat.validate_user_exists(1, mock_db) is None

# 유저가 존재하지 않을 때 예외가 발생하는지 테스트
def test_validate_user_exists_fail():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None
    with pytest.raises(UserNotFoundException):
        chat.validate_user_exists(1, mock_db)

# touch_session이 세션의 수정시간을 현재로 변경하는지 테스트
def test_touch_session_updates_modifiedAt():
    mock_db = MagicMock()
    before = datetime(2020, 1, 1, tzinfo=KST)  # ✅ 타임존-aware로 변경
    session = ChatSession(id=1, userId=1, name="Test", summary="", modifiedAt=before)
    mock_db.query().filter().first.return_value = session

    chat.touch_session(sessionId=1, db=mock_db)

    assert session.modifiedAt > before  # ✅ now(KST) > before(KST)

# 세션 요약(summary)을 정상적으로 업데이트하는지 테스트
def test_update_session_summary_success():
    mock_db = MagicMock()
    session = ChatSession(id=1, userId=1, name="Test", summary="old")
    mock_db.query().filter().first.return_value = session
    chat.update_session_summary(sessionId=1, summary="new summary", db=mock_db)
    assert session.summary == "new summary"

# 세션이 없을 경우 summary 업데이트 시 예외가 발생하는지 테스트
def test_update_session_summary_fail():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None
    with pytest.raises(SessionNotFoundException):
        chat.update_session_summary(1, "test", mock_db)

# ChatSheet가 없을 경우 새로 삽입하는지 테스트
def test_upsert_chat_sheet_insert():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None
    result = chat.upsert_chat_sheet(1, b"hello", mock_db)
    assert result.sheetData == b"hello"

# 기존 ChatSheet가 존재할 경우 sheetData를 업데이트하는지 테스트
def test_upsert_chat_sheet_update():
    mock_db = MagicMock()
    existing_sheet = ChatSheet(sessionId=1, sheetData=b"old")
    mock_db.query().filter().first.return_value = existing_sheet
    result = chat.upsert_chat_sheet(1, b"new", mock_db)
    assert result.sheetData == b"new"

# 메시지를 데이터베이스에 정상 삽입하는지 테스트
def test_insert_message_to_db():
    mock_db = MagicMock()
    mock_db.flush = MagicMock()
    mock_db.add = MagicMock()
    result = chat.insert_message_to_db(1, "hello", "USER", mock_db)
    assert result.content == "hello"
    assert result.senderType == "USER"


def test_delete_session_success():
    mock_db = MagicMock()
    session = ChatSession(id=1, userId=1)
    mock_db.query().filter().first.return_value = session

    chat.delete_session(1, mock_db)

    mock_db.delete.assert_called_once_with(session)
    mock_db.commit.assert_called_once()