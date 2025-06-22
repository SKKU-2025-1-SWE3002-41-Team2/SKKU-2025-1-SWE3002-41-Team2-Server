import pytest
from unittest.mock import MagicMock
from app.services.auth_service import login
from app.models import User
from app.exceptions.http_exceptions import UnauthorizedException


def test_login_success():
    # Arrange
    mock_db = MagicMock()
    mock_user = User(id=1, username="testuser", password="testpass")
    mock_db.query().filter().first.return_value = mock_user

    # Act
    result = login(mock_db, "testuser", "testpass")

    # Assert
    assert result == mock_user


def test_login_user_not_found():
    mock_db = MagicMock()
    mock_db.query().filter().first.return_value = None

    with pytest.raises(UnauthorizedException):
        login(mock_db, "wronguser", "any")


def test_login_wrong_password():
    mock_db = MagicMock()
    mock_user = User(id=1, username="testuser", password="correctpass")
    mock_db.query().filter().first.return_value = mock_user

    with pytest.raises(UnauthorizedException):
        login(mock_db, "testuser", "wrongpass")
