"""Agent-facing terminal tools.

Everything in this package is **read-only** by construction: agents explore
the hospital database through these tools without being able to mutate state.
See ``docs/agent_interface.md`` for the full agent guide.
"""

from setup.agent.sql_tool import QueryResult, SQLGuardError, run_sql, validate_readonly_sql

__all__ = ["QueryResult", "SQLGuardError", "run_sql", "validate_readonly_sql"]
