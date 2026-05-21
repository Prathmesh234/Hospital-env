"""Typer CLI entry point: ``uv run hospital-env <command>``."""

from __future__ import annotations

from pathlib import Path

import typer
import uvicorn
from rich.console import Console
from rich.table import Table
from sqlalchemy import text

from setup.config import get_settings
from setup.db.mongo import get_sync_db
from setup.db.postgres import get_engine, get_session_factory
from setup.ingest import diff_workbook, list_expected_sheets, load_workbook
from setup.models import Base
from setup.nosql import drop_collections, ensure_collections

app = typer.Typer(
    help="Hospital-env CLI — schema setup, data loading, and API server.",
    no_args_is_help=True,
)
console = Console()


@app.command("init-db")
def init_db(
    mongo: bool = typer.Option(True, help="Also bootstrap MongoDB collections + indexes."),
) -> None:
    """Create the full PostgreSQL schema (DDL only) and Mongo collections."""
    engine = get_engine()
    console.print(f"[green]Creating {len(Base.metadata.tables)} tables in PostgreSQL…[/green]")
    Base.metadata.create_all(engine)
    console.print("[green]✓ PostgreSQL schema ready.[/green]")

    if mongo:
        db = get_sync_db()
        created = ensure_collections(db)
        if created:
            console.print(f"[green]✓ Created Mongo collections: {', '.join(created)}[/green]")
        else:
            console.print("[green]✓ Mongo collections already present (indexes refreshed).[/green]")


@app.command("drop-db")
def drop_db(
    confirm: bool = typer.Option(False, "--yes", "-y", help="Skip confirmation prompt."),
    mongo: bool = typer.Option(True, help="Also drop MongoDB collections."),
) -> None:
    """Drop the entire schema (DESTRUCTIVE)."""
    if not confirm:
        typer.confirm("Are you sure? This will drop every table.", abort=True)
    engine = get_engine()
    Base.metadata.drop_all(engine)
    console.print("[red]✓ PostgreSQL schema dropped.[/red]")
    if mongo:
        dropped = drop_collections(get_sync_db())
        if dropped:
            console.print(f"[red]✓ Dropped Mongo collections: {', '.join(dropped)}[/red]")


@app.command()
def load(
    file: Path = typer.Argument(..., exists=True, readable=True, help="Path to .xlsx workbook."),
) -> None:
    """Load a multi-sheet xlsx workbook into the hospital schema."""
    factory = get_session_factory()
    with factory() as session:
        counts = load_workbook(session, file)

    table = Table(title=f"Loaded {file.name}", show_lines=False)
    table.add_column("Table")
    table.add_column("Rows", justify="right")
    for name, count in counts.items():
        table.add_row(name, str(count))
    console.print(table)
    console.print(f"[green]✓ Loaded {sum(counts.values())} rows across {len(counts)} tables.[/green]")


@app.command("diff-workbook")
def diff_workbook_cmd(
    file: Path = typer.Argument(..., exists=True, readable=True),
) -> None:
    """Compare a workbook against the canonical sheet list."""
    diff = diff_workbook(file)
    console.print("[bold]Missing sheets[/bold]:")
    for name in diff["missing_sheets"]:
        console.print(f"  • {name}")
    console.print("[bold]Extra sheets[/bold]:")
    for name in diff["extra_sheets"]:
        console.print(f"  • {name}")


@app.command("list-sheets")
def list_sheets() -> None:
    """Print the canonical sheet list (load order)."""
    for i, name in enumerate(list_expected_sheets(), start=1):
        console.print(f"{i:>3}. {name}")


@app.command()
def serve(
    host: str | None = typer.Option(None, help="Override API_HOST."),
    port: int | None = typer.Option(None, help="Override API_PORT."),
    reload: bool | None = typer.Option(None, help="Override API_RELOAD."),
) -> None:
    """Run the FastAPI server with uvicorn."""
    settings = get_settings()
    uvicorn.run(
        "setup.api.main:app",
        host=host or settings.api_host,
        port=port or settings.api_port,
        reload=settings.api_reload if reload is None else reload,
        log_level=settings.log_level.lower(),
    )


@app.command()
def status() -> None:
    """Quick health check for PostgreSQL + MongoDB."""
    settings = get_settings()
    console.print(f"[bold]Postgres URL[/bold]: {settings.sqlalchemy_url}")
    console.print(f"[bold]Mongo URI   [/bold]: {settings.mongo_uri}")

    try:
        with get_engine().connect() as conn:
            conn.execute(text("SELECT 1"))
        console.print("[green]✓ PostgreSQL reachable[/green]")
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]✗ PostgreSQL error: {exc}[/red]")

    try:
        get_sync_db().command("ping")
        console.print("[green]✓ MongoDB reachable[/green]")
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]✗ MongoDB error: {exc}[/red]")


if __name__ == "__main__":
    app()
