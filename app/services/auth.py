from sqlalchemy.orm import Session
from app.exceptions.http_exceptions import UnauthorizedException
from app.models import User

"""
Interface Summary:
- def login(db: Session, username: str, password: str) -> User
"""


def login(db: Session, username: str, password: str):
    user = db.query(User).filter(User.username == username).first()

    if not user or user.password != password:
        raise UnauthorizedException()
    return user