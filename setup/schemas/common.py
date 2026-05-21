"""Shared Pydantic base config for all schemas."""

from __future__ import annotations

from pydantic import BaseModel, ConfigDict


class ORMModel(BaseModel):
    """Base for read schemas — enables loading directly from ORM instances."""

    model_config = ConfigDict(from_attributes=True, populate_by_name=True)
