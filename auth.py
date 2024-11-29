from datetime import datetime, timedelta, timezone
from typing import Optional, Dict, Any, cast
from jose import JWTError, jwt
from passlib.context import CryptContext
from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from pydantic import BaseModel
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

SECRET_KEY = os.getenv("SECRET_KEY")
if not SECRET_KEY:
    raise ValueError("SECRET_KEY not found in environment variables")

ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 30


class User(BaseModel):
    """Model for user information."""

    username: str
    disabled: Optional[bool] = None


class UserInDB(User):
    """Model for user information stored in the database."""

    hashed_password: str


class Token(BaseModel):
    """Model for the access token response."""

    access_token: str
    token_type: str


class TokenData(BaseModel):
    """Model for token data."""

    username: Optional[str] = None


# Password hashing context
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")

# TODO: Replace with a real database
users_db = {
    "admin": {
        "username": "admin",
        "hashed_password": pwd_context.hash("admin123"),  # Change in production
        "disabled": False,
    },
    "user": {
        "username": "user",
        "hashed_password": pwd_context.hash("user123"),  # Change in production
        "disabled": False,
    },
}


def verify_password(plain_password: str, hashed_password: str) -> bool:
    """Verify a plain password against a hashed password.

    Args:
        plain_password (str): The plain password to verify.
        hashed_password (str): The hashed password to compare against.

    Returns:
        bool: True if the password matches, False otherwise.
    """
    return pwd_context.verify(plain_password, hashed_password)


def get_user(db, username: str) -> Optional[UserInDB]:
    """Retrieve a user from the database by username.

    Args:
        db: The database of users.
        username (str): The username of the user to retrieve.

    Returns:
        UserInDB: The user object if found, None otherwise.
    """
    if username in db:
        user_dict = db[username]
        return UserInDB(**user_dict)


def authenticate_user(db, username: str, password: str) -> Optional[UserInDB]:
    """Authenticate a user by username and password.

    Args:
        db: The database of users.
        username (str): The username of the user.
        password (str): The password of the user.

    Returns:
        UserInDB: The authenticated user object if successful, None otherwise.
    """
    user = get_user(db, username)
    if not user:
        return None
    if not verify_password(password, user.hashed_password):
        return None
    return user


def create_access_token(
    data: Dict[str, Any], expires_delta: Optional[timedelta] = None
) -> str:
    """Create a new access token.

    Args:
        data (Dict[str, Any]): The data to encode in the token.
        expires_delta (Optional[timedelta]): The expiration time for the token.

    Returns:
        str: The encoded JWT access token.
    """
    to_encode = data.copy()

    if expires_delta:
        expire = datetime.now(timezone.utc) + expires_delta
    else:
        expire = datetime.now(timezone.utc) + timedelta(
            minutes=ACCESS_TOKEN_EXPIRE_MINUTES
        )

    to_encode.update({"exp": expire})
    encoded_jwt = jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)
    return encoded_jwt


async def get_current_user(token: str = Depends(oauth2_scheme)) -> User:
    """Retrieve the current user based on the provided token.

    Args:
        token (str): The JWT token provided by the user.

    Returns:
        User: The current authenticated user.

    Raises:
        HTTPException: If the token is invalid or the user is not found.
    """
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Invalid credentials",
        headers={"WWW-Authenticate": "Bearer"},
    )

    try:
        # Explicitly type the payload
        payload: Dict[str, Any] = jwt.decode(
            token=token, key=SECRET_KEY, algorithms=[ALGORITHM]
        )

        # Extract and type the username
        username: Optional[str] = payload.get("sub")
        if username is None:
            raise credentials_exception

        # Create token data with explicit typing
        token_data = TokenData(username=username)

    except JWTError as e:
        raise credentials_exception from e

    # Get user with explicit return type checking
    user: Optional[UserInDB] = get_user(users_db, cast(str, token_data.username))
    if user is None:
        raise credentials_exception

    return user
