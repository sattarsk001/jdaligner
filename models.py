from sqlalchemy import Column, Integer, String, DateTime, func, ForeignKey
from database import Base

class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String, unique=True, index=True, nullable=False)
    password_hash = Column(String, nullable=True)  # for email/password users
    google_id = Column(String, nullable=True)      # for Google OAuth users
    created_at = Column(DateTime, server_default=func.now())
    updated_at = Column(
        DateTime, server_default=func.now(), onupdate=func.now()
    )


class PasswordResetToken(Base):
    __tablename__ = "password_reset_tokens"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)

    # hashed version of the reset token we email to user
    token_hash = Column(String, nullable=False)

    # when this token expires
    expires_at = Column(DateTime, nullable=False)

    # when the token was actually used (null = not used yet)
    used_at = Column(DateTime, nullable=True)

    created_at = Column(DateTime, server_default=func.now())
