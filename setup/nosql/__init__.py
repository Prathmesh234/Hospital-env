"""MongoDB collections used as document-shaped side-stores."""

from setup.nosql.collections import COLLECTIONS, drop_collections, ensure_collections

__all__ = ["COLLECTIONS", "ensure_collections", "drop_collections"]
