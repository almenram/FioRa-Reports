import sys
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, sessionmaker


if getattr(sys, "frozen", False):
    # Cuando corre como .exe
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    # Cuando corre como .py
    BASE_DIR = Path(__file__).resolve().parent


DATABASE_URL = f"sqlite:///{BASE_DIR / 'reportes.db'}"


engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False}
)


SessionLocal = sessionmaker(
    bind=engine,
    autoflush=False,
    expire_on_commit=False
)


class Base(DeclarativeBase):
    pass