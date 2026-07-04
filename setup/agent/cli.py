"""`hospital-env agent …` — the terminal tool surface for agents.

Commands:
    agent sql       run one guarded read-only SQL query
    agent tables    list tables (domain + live row count)
    agent describe  column/FK/index detail for one table
    agent sample    peek at a few rows of a table
    agent mongo     read-only find/count on the document store
    agent api       GET an endpoint of the running REST API
    agent shell     interactive REPL combining all of the above
"""

from __future__ import annotations

import json

import typer
from rich.console import Console
from rich.table import Table

from setup.agent import api_tool, inspect_tool, mongo_tool
from setup.agent.sql_tool import (
    DEFAULT_ROW_LIMIT,
    DEFAULT_TIMEOUT_MS,
    QueryResult,
    SQLGuardError,
    run_sql,
)

agent_app = typer.Typer(
    help="Read-only agent tools: SQL, schema inspection, Mongo, REST API.",
    no_args_is_help=True,
)
console = Console()


def _render(result: QueryResult, fmt: str) -> None:
    if fmt == "json":
        console.print_json(result.to_json())
        return
    if fmt == "csv":
        typer.echo(result.to_csv(), nl=False)
        return
    table = Table(show_lines=False)
    for col in result.columns:
        table.add_column(col, overflow="fold")
    for row in result.rows:
        table.add_row(*["" if v is None else str(v) for v in row])
    console.print(table)
    suffix = " (truncated — raise --limit)" if result.truncated else ""
    console.print(f"[dim]{result.rowcount} row(s) in {result.elapsed_ms:.1f} ms{suffix}[/dim]")


@agent_app.command("sql")
def sql_cmd(
    query: str = typer.Argument(..., help="A single SELECT/WITH query."),
    limit: int = typer.Option(DEFAULT_ROW_LIMIT, help="Max rows returned (cap 1000)."),
    timeout_ms: int = typer.Option(DEFAULT_TIMEOUT_MS, help="Server-side statement timeout."),
    fmt: str = typer.Option("table", "--format", "-f", help="table | json | csv"),
) -> None:
    """Run one read-only SQL query against PostgreSQL."""
    try:
        result = run_sql(query, limit=limit, timeout_ms=timeout_ms)
    except SQLGuardError as exc:
        console.print(f"[red]rejected:[/red] {exc}")
        raise typer.Exit(code=2) from exc
    except Exception as exc:  # noqa: BLE001 — show the DB error, not a traceback
        message = getattr(getattr(exc, "orig", None), "args", None) or (str(exc),)
        console.print(f"[red]sql error:[/red] {message[0]}")
        raise typer.Exit(code=1) from exc
    _render(result, fmt)


@agent_app.command("tables")
def tables_cmd(
    domain: str | None = typer.Option(None, help="Filter to one domain."),
    no_counts: bool = typer.Option(False, "--no-counts", help="Skip live row counts."),
) -> None:
    """List every table with its domain and current row count."""
    rows = inspect_tool.list_tables(domain=domain, with_counts=not no_counts)
    if not rows:
        console.print(
            f"[yellow]No tables in domain {domain!r}. "
            f"Domains: {', '.join(inspect_tool.list_domains())}[/yellow]"
        )
        raise typer.Exit(code=1)
    table = Table(title=f"{len(rows)} tables")
    table.add_column("domain")
    table.add_column("table")
    table.add_column("rows", justify="right")
    for row in rows:
        table.add_row(row["domain"], row["table"], "" if row["rows"] is None else str(row["rows"]))
    console.print(table)


@agent_app.command("describe")
def describe_cmd(table_name: str = typer.Argument(..., help="Table to describe.")) -> None:
    """Show columns, keys, indexes, and referencing tables for one table."""
    try:
        info = inspect_tool.describe_table(table_name)
    except KeyError as exc:
        console.print(f"[red]{exc.args[0]}[/red]")
        raise typer.Exit(code=1) from exc

    console.print(f"[bold]{info['table']}[/bold]  (domain: {info['domain']})")
    table = Table(show_lines=False)
    for header in ("column", "type", "null", "key"):
        table.add_column(header)
    for col in info["columns"]:
        key = "PK" if col.primary_key else ""
        if col.foreign_keys:
            key = (key + " " if key else "") + "→ " + ", ".join(col.foreign_keys)
        table.add_row(col.name, col.type, "" if col.nullable else "NOT NULL", key)
    console.print(table)
    if info["indexes"]:
        console.print("[bold]indexes[/bold]")
        for idx in info["indexes"]:
            uniq = " UNIQUE" if idx["unique"] else ""
            console.print(f"  • {idx['name']}{uniq} ({', '.join(idx['columns'])})")
    if info["referenced_by"]:
        console.print(f"[bold]referenced by[/bold]: {', '.join(info['referenced_by'])}")


@agent_app.command("sample")
def sample_cmd(
    table_name: str = typer.Argument(...),
    limit: int = typer.Option(5, help="Rows to fetch."),
    fmt: str = typer.Option("table", "--format", "-f", help="table | json | csv"),
) -> None:
    """Peek at a few rows of one table."""
    try:
        result = inspect_tool.sample_table(table_name, limit=limit)
    except KeyError as exc:
        console.print(f"[red]{exc.args[0]}[/red]")
        raise typer.Exit(code=1) from exc
    _render(result, fmt)


@agent_app.command("mongo")
def mongo_cmd(
    collection: str = typer.Argument(
        None,
        help="Collection name; omit to list collections with document counts.",
    ),
    filter_json: str = typer.Option("{}", "--filter", help="Mongo filter as JSON."),
    projection_json: str = typer.Option(None, "--project", help="Projection as JSON."),
    sort: str = typer.Option(None, help="Sort as 'field:asc' or 'field:desc'."),
    limit: int = typer.Option(mongo_tool.DEFAULT_LIMIT, help="Max documents (cap 200)."),
    count_only: bool = typer.Option(False, "--count", help="Return the count only."),
) -> None:
    """Read-only find/count on the MongoDB document store."""
    if collection is None:
        table = Table(title="Mongo collections")
        table.add_column("collection")
        table.add_column("documents", justify="right")
        for row in mongo_tool.collection_stats():
            table.add_row(row["collection"], str(row["documents"]))
        console.print(table)
        return

    try:
        filter_ = mongo_tool.parse_json_arg(filter_json, "--filter") or {}
        projection = mongo_tool.parse_json_arg(projection_json, "--project")
        if count_only:
            console.print(str(mongo_tool.count(collection, filter_)))
            return
        sort_spec = None
        if sort:
            field, _, direction = sort.partition(":")
            sort_spec = [(field, -1 if direction.lower() == "desc" else 1)]
        docs = mongo_tool.find(collection, filter_, projection, sort_spec, limit)
    except mongo_tool.MongoGuardError as exc:
        console.print(f"[red]rejected:[/red] {exc}")
        raise typer.Exit(code=2) from exc
    console.print_json(json.dumps(docs, default=str))


@agent_app.command("api")
def api_cmd(
    path: str = typer.Argument(..., help="Endpoint path, e.g. /patients"),
    param: list[str] = typer.Option(None, "--param", "-p", help="Query param key=value (repeatable)."),
) -> None:
    """GET an endpoint of the running REST API (see /docs for the catalog)."""
    params: dict[str, str] = {}
    for item in param or []:
        key, sep, value = item.partition("=")
        if not sep:
            console.print(f"[red]Bad --param {item!r}; expected key=value.[/red]")
            raise typer.Exit(code=2)
        params[key] = value
    try:
        status, body = api_tool.get(path, params)
    except api_tool.APIToolError as exc:
        console.print(f"[red]{exc}[/red]")
        raise typer.Exit(code=1) from exc
    console.print(f"[dim]HTTP {status}[/dim]")
    if isinstance(body, (dict, list)):
        console.print_json(json.dumps(body, default=str))
    else:
        typer.echo(body)


_SHELL_HELP = """\
Interactive agent shell. End SQL with ';' to execute. Meta commands:
  \\t [domain]         list tables (optionally one domain)
  \\d <table>          describe a table
  \\s <table> [n]      sample n rows (default 5)
  \\m [coll] [filter]  mongo: list collections, or find with a JSON filter
  \\api <path>         GET a REST endpoint
  \\f table|json|csv   switch SQL output format
  \\h                  this help
  \\q                  quit
"""


@agent_app.command("shell")
def shell_cmd() -> None:
    """Interactive REPL over all agent tools."""
    console.print("[bold green]hospital-env agent shell[/bold green] — \\h for help, \\q to quit")
    fmt = "table"
    buffer: list[str] = []
    while True:
        try:
            prompt = "sql> " if not buffer else "...> "
            line = input(prompt)
        except (EOFError, KeyboardInterrupt):
            console.print()
            break

        stripped = line.strip()
        if not buffer and stripped.startswith("\\"):
            parts = stripped.split()
            cmd, args = parts[0], parts[1:]
            try:
                if cmd == "\\q":
                    break
                elif cmd == "\\h":
                    typer.echo(_SHELL_HELP)
                elif cmd == "\\t":
                    tables_cmd(domain=args[0] if args else None, no_counts=False)
                elif cmd == "\\d" and args:
                    describe_cmd(args[0])
                elif cmd == "\\s" and args:
                    sample_cmd(args[0], limit=int(args[1]) if len(args) > 1 else 5, fmt=fmt)
                elif cmd == "\\m":
                    if not args:
                        mongo_cmd(collection=None)
                    else:
                        mongo_cmd(
                            collection=args[0],
                            filter_json=" ".join(args[1:]) or "{}",
                            projection_json=None,
                            sort=None,
                            limit=mongo_tool.DEFAULT_LIMIT,
                            count_only=False,
                        )
                elif cmd == "\\api" and args:
                    api_cmd(args[0], param=[])
                elif cmd == "\\f" and args:
                    if args[0] in {"table", "json", "csv"}:
                        fmt = args[0]
                        console.print(f"[dim]format = {fmt}[/dim]")
                    else:
                        console.print("[red]format must be table, json, or csv[/red]")
                else:
                    console.print("[red]Unknown meta command — \\h for help.[/red]")
            except typer.Exit:
                pass
            continue

        buffer.append(line)
        if stripped.endswith(";"):
            query = "\n".join(buffer)
            buffer = []
            try:
                _render(run_sql(query), fmt)
            except SQLGuardError as exc:
                console.print(f"[red]rejected:[/red] {exc}")
            except Exception as exc:  # noqa: BLE001 — surface DB errors, keep the REPL alive
                console.print(f"[red]error:[/red] {exc}")
