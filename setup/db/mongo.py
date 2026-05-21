"""MongoDB clients (sync PyMongo + async Motor)."""

from __future__ import annotations

from motor.motor_asyncio import AsyncIOMotorClient, AsyncIOMotorDatabase
from pymongo import MongoClient
from pymongo.database import Database

from setup.config import get_settings

_sync_client: MongoClient | None = None
_async_client: AsyncIOMotorClient | None = None


def get_sync_client() -> MongoClient:
    global _sync_client
    if _sync_client is None:
        _sync_client = MongoClient(get_settings().mongo_uri)
    return _sync_client


def get_sync_db() -> Database:
    return get_sync_client()[get_settings().mongo_db]


def get_async_client() -> AsyncIOMotorClient:
    global _async_client
    if _async_client is None:
        _async_client = AsyncIOMotorClient(get_settings().mongo_uri)
    return _async_client


def get_async_db() -> AsyncIOMotorDatabase:
    return get_async_client()[get_settings().mongo_db]
