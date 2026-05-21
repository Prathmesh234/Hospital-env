"""PostgreSQL engine + session factory."""

from __future__ import annotations

from collections.abc import Generator

from sqlalchemy import create_engine
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session, sessionmaker

from setup.config import get_settings


def make_engine(echo: bool = False) -> Engine:
    settings = get_settings()
    return create_engine(
        settings.sqlalchemy_url,
        echo=echo,
        pool_pre_ping=True,
        future=True,
    )


_engine: Engine | None = None
_SessionLocal: sessionmaker[Session] | None = None


def get_engine() -> Engine:
    global _engine, _SessionLocal
    if _engine is None:
        _engine = make_engine()
        _SessionLocal = sessionmaker(bind=_engine, autoflush=False, expire_on_commit=False)
    return _engine


def get_session_factory() -> sessionmaker[Session]:
    get_engine()
    assert _SessionLocal is not None
    return _SessionLocal


def get_session() -> Generator[Session, None, None]:
    """FastAPI dependency yielding a SQLAlchemy session."""
    factory = get_session_factory()
    session = factory()
    try:
        yield session
    finally:
        session.close()
