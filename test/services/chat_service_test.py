import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime
from app.services import chat_service
from app.models import ChatSession, Message, ChatSheet, User
from app.exceptions.http_exceptions import SessionNotFoundException, UserNotFoundException
from app.utils.timezone import KST
from app.schemas.chat_schema import ChatSessionCreateResponse, LLMMessageResponse, MessageResponse

# [GET] 사용자의 세션 목록을 정상적으로 불러올 수 있는지 테스트
def test_get_sessions_success():
    mock_db = MagicMock()
    mock_db.query().filter().order_by().all.return_value = [
        ChatSession(id=1, userId=10, name="Test Session", summary="test", modifiedAt=datetime.now())
    ]
    sessions = chat_service.get_sessions(userId=10, db=mock_db)
    assert len(sessions) == 1
    assert sessions[0].userId == 10

# [GET] 세션이 없을 때 예외가 발생하는지 테스트
def test_get_sessions_not_found():
    mock_db = MagicMock()
    mock_db.query().filter().order_by().all.return_value = []
    with pytest.raises(SessionNotFoundException):
        chat_service.get_sessions(userId=99, db=mock_db)

# [GET] 특정 세션의 메시지를 정상 조회할 수 있는지 테스트
def test_get_messages_success():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = ChatSession(id=1, userId=1)
    session = chat_service.get_messages(session_id=1, db=mock_db)
    assert session.id == 1

# [GET] 존재하지 않는 세션 조회 시 예외 발생하는지 테스트
def test_get_messages_not_found():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None
    with pytest.raises(SessionNotFoundException):
        chat_service.get_messages(session_id=1, db=mock_db)

# [AUTH] 유저가 존재할 때 정상적으로 통과되는지 테스트
def test_validate_user_exists_success():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = User(id=1)
    assert chat_service.validate_user_exists(1, mock_db) is None

# [AUTH] 유저가 존재하지 않을 때 예외가 발생하는지 테스트
def test_validate_user_exists_fail():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None
    with pytest.raises(UserNotFoundException):
        chat_service.validate_user_exists(1, mock_db)

# [TOUCH] touch_session이 세션의 수정시간을 현재로 변경하는지 테스트
def test_touch_session_updates_modifiedAt():
    mock_db = MagicMock()
    before = datetime(2020, 1, 1, tzinfo=KST)  # ✅ 타임존-aware로 변경
    session = ChatSession(id=1, userId=1, name="Test", summary="", modifiedAt=before)
    mock_db.query().filter().first.return_value = session

    chat_service.touch_session(sessionId=1, db=mock_db)

    assert session.modifiedAt > before  # ✅ now(KST) > before(KST)

# [UPDATE] 세션 요약(summary)을 정상적으로 업데이트하는지 테스트
def test_update_session_summary_success():
    mock_db = MagicMock()
    session = ChatSession(id=1, userId=1, name="Test", summary="old")
    mock_db.query().filter().first.return_value = session
    chat_service.update_session_summary(sessionId=1, summary="new summary", db=mock_db)
    assert session.summary == "new summary"

# [UPDATE] 세션이 없을 경우 summary 업데이트 시 예외가 발생하는지 테스트
def test_update_session_summary_fail():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None
    with pytest.raises(SessionNotFoundException):
        chat_service.update_session_summary(1, "test", mock_db)

# [UPSERT] ChatSheet가 없을 경우 새로 삽입하는지 테스트
def test_upsert_chat_sheet_insert():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None
    result = chat_service.upsert_chat_sheet(1, b"hello", mock_db)
    assert result.sheetData == b"hello"

# [UPSERT] 기존 ChatSheet가 존재할 경우 sheetData를 업데이트하는지 테스트
def test_upsert_chat_sheet_update():
    mock_db = MagicMock()
    existing_sheet = ChatSheet(sessionId=1, sheetData=b"old")
    mock_db.query().filter().first.return_value = existing_sheet
    result = chat_service.upsert_chat_sheet(1, b"new", mock_db)
    assert result.sheetData == b"new"

# [INSERT] 메시지를 데이터베이스에 정상 삽입하는지 테스트
def test_insert_message_to_db():
    mock_db = MagicMock()
    mock_db.flush = MagicMock()
    mock_db.add = MagicMock()
    result = chat_service.insert_message_to_db(1, "hello", "USER", mock_db)
    assert result.content == "hello"
    assert result.senderType == "USER"

# [DELETE] 세션이 존재할 경우 정상적으로 삭제되는지 테스트
def test_delete_session_success():
    mock_db = MagicMock()
    session = ChatSession(id=1, userId=1)
    mock_db.query().filter().first.return_value = session

    chat_service.delete_session(1, mock_db)

    mock_db.delete.assert_called_once_with(session)
    mock_db.commit.assert_called_once()


# [DELETE] 세션이 존재하지 않을 경우 예외가 발생하는지 테스트
def test_delete_session_not_found():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None

    with pytest.raises(SessionNotFoundException):
        chat_service.delete_session(1, mock_db)


# [MODIFY] 세션 이름이 정상적으로 변경되는지 테스트
def test_modify_session_success():
    mock_db = MagicMock()
    session = ChatSession(id=1, userId=1, name="Old Name")
    mock_db.query().filter().first.return_value = session

    result = chat_service.modify_session(1, "New Name", mock_db)

    assert result.name == "New Name"
    mock_db.commit.assert_called_once()
    mock_db.refresh.assert_called_once_with(session)


# [MODIFY] 세션이 존재하지 않을 경우 예외가 발생하는지 테스트
def test_modify_session_not_found():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None

    with pytest.raises(SessionNotFoundException):
        chat_service.modify_session(1, "New Name", mock_db)


# [CREATE] create_session 호출 시 내부 흐름이 정상적으로 작동하는지 테스트
@patch("app.services.chat_service.validate_user_exists")
@patch("app.services.chat_service.save_message_and_response")
def test_create_session_flow(mock_save_msg, mock_validate_user):
    mock_db = MagicMock()

    # 👉 ChatSession 객체가 생성될 때 session.id = 1로 지정
    fake_session = ChatSession(userId=1, name="New Session")
    fake_session.id = 123  # ✅ 직접 ID 지정
    mock_db.flush = MagicMock()
    mock_db.add = MagicMock()

    # save_message_and_response 결과도 더미 반환
    mock_save_msg.return_value = LLMMessageResponse(
        sheetData="base64data",
        message=MessageResponse(id=1, content="reply", createdAt=datetime.now(), senderType="AI")
    )

    # 👉 patch가 아닌 내부에서 사용될 세션 생성 시 return 값을 조작
    with patch("app.services.chat_service.ChatSession", return_value=fake_session):
        result = chat_service.create_session(1, "Hello", b"excel-bytes", mock_db)

    # 검증
    mock_validate_user.assert_called_once_with(1, mock_db)
    mock_save_msg.assert_called_once()
    assert isinstance(result, ChatSessionCreateResponse)
    assert result.sessionId == 123
    assert result.message.content == "reply"


# [SAVE] save_message_and_response에서 모든 의존 함수가 호출되는지 테스트
@patch("app.services.chat_service.insert_message_to_db")
@patch("app.services.chat_service.get_llm_response")
@patch("app.services.chat_service.process_excel_with_commands")
@patch("app.services.chat_service.update_session_summary")
@patch("app.services.chat_service.upsert_chat_sheet")
def test_save_message_and_response_flow(
    mock_upsert,
    mock_update_summary,
    mock_process_excel,
    mock_get_llm,
    mock_insert_msg
):
    mock_db = MagicMock()
    session = ChatSession(id=1, userId=1, summary="prev-summary")
    mock_db.query().filter().first.return_value = session

    # 첫 번째는 USER 메시지, 두 번째는 AI 메시지
    mock_insert_msg.side_effect = [
        MagicMock(id=10, content="user-message", createdAt=datetime.now(), senderType="USER"),
        MagicMock(id=11, content="ai-reply", createdAt=datetime.now(), senderType="AI"),
    ]

    mock_get_llm.return_value = MagicMock(
        chat="ai-reply",
        summary="updated-summary",
        cmd_seq=[{"command_type": "sum"}]
    )
    mock_process_excel.return_value = b"new-excel-bytes"

    result = chat_service.save_message_and_response(1, "Hi", b"old-bytes", mock_db)

    assert isinstance(result, LLMMessageResponse)
    assert result.message.content == "ai-reply"
    mock_insert_msg.assert_called()
    mock_get_llm.assert_called_once()
    mock_process_excel.assert_called_once()
    mock_update_summary.assert_called_once_with(sessionId=1, summary="updated-summary", db=mock_db)
    mock_upsert.assert_called_once_with(1, b"new-excel-bytes", mock_db)
    mock_db.commit.assert_called_once()