from sqlalchemy.orm import Session
from app.models import User
from fastapi import HTTPException, status

"""
Interface Summary:
- def login(db: Session, username: str, password: str) -> User
"""


def login(db: Session, username: str, password: str):
    user = db.query(User).filter(User.username == username).first()
    print("gooood")
    print(user)

    if not user or user.password != password:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="아이디 또는 비밀번호가 틀립니다"
        )
    return user