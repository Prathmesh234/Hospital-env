"""Centralized settings loaded from environment / `.env`."""

from __future__ import annotations

from functools import lru_cache

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings.

    Values are read from environment variables and, optionally, a `.env` file
    in the project root. See `.env.example` for the full list.
    """

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # PostgreSQL
    postgres_host: str = "localhost"
    postgres_port: int = 5432
    postgres_db: str = "hospital"
    postgres_user: str = "hospital"
    postgres_password: str = "hospital_dev_password"
    postgres_url: str | None = None

    # MongoDB
    mongo_host: str = "localhost"
    mongo_port: int = 27017
    mongo_db: str = "hospital_docs"
    mongo_user: str = "hospital"
    mongo_password: str = "hospital_dev_password"
    mongo_url: str | None = None

    # API
    api_host: str = "0.0.0.0"
    api_port: int = 8000
    api_reload: bool = True
    log_level: str = Field(default="INFO")

    @property
    def sqlalchemy_url(self) -> str:
        if self.postgres_url:
            return self.postgres_url
        return (
            f"postgresql+psycopg://{self.postgres_user}:{self.postgres_password}"
            f"@{self.postgres_host}:{self.postgres_port}/{self.postgres_db}"
        )

    @property
    def mongo_uri(self) -> str:
        if self.mongo_url:
            return self.mongo_url
        return (
            f"mongodb://{self.mongo_user}:{self.mongo_password}"
            f"@{self.mongo_host}:{self.mongo_port}/{self.mongo_db}?authSource=admin"
        )


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()
