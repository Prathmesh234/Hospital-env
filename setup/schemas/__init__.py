"""Pydantic v2 schemas mirroring the ORM models.

These are lightweight pass-through schemas for the FastAPI surface.  Each
domain module exposes a ``Read`` schema (everything the DB knows) and a
``Create`` schema (the subset a client can submit).
"""

from setup.schemas.common import ORMModel

__all__ = ["ORMModel"]
