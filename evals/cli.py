"""`hospital-env task …` — hand questions to agents and grade their answers.

The intended loop:

    hospital-env task show HOSP-016      # the query sent to the agent
    …agent works the terminal tools…
    hospital-env task submit HOSP-016 --answer "59.7"

`task list`/`task show` never print gold answers. Agents are expected NOT to
read evals/tasks.jsonl — for hard isolation, hand out the file produced by
`task export-questions` instead.
"""

from __future__ import annotations

import json
from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from evals import runner
from evals.grader import grade

task_app = typer.Typer(
    help="Evaluation tasks: list/show questions, submit answers, verify golds.",
    no_args_is_help=True,
)
console = Console()


@task_app.command("list")
def list_cmd(
    category: str | None = typer.Option(None, help="Filter by category."),
    difficulty: str | None = typer.Option(None, help="easy | medium | hard"),
) -> None:
    """List all task ids (questions truncated; never shows answers)."""
    tasks = runner.load_tasks()
    table = Table(title=f"{len(tasks)} evaluation tasks")
    for header in ("id", "category", "difficulty", "type", "question"):
        table.add_column(header, overflow="fold")
    shown = 0
    for task in tasks:
        if category and task["category"] != category:
            continue
        if difficulty and task["difficulty"] != difficulty:
            continue
        question = task["question"]
        table.add_row(
            task["id"],
            task["category"],
            task["difficulty"],
            task["answer_type"],
            question if len(question) <= 88 else question[:85] + "…",
        )
        shown += 1
    console.print(table)
    if shown == 0:
        console.print("[yellow]No tasks matched the filters.[/yellow]")


@task_app.command("show")
def show_cmd(task_id: str = typer.Argument(..., help="e.g. HOSP-016")) -> None:
    """Print one task's question — this is the query an agent receives."""
    try:
        task = runner.get_task(task_id)
    except KeyError as exc:
        console.print(f"[red]{exc.args[0]}[/red]")
        raise typer.Exit(code=1) from exc
    view = runner.public_view(task)
    console.print(f"[bold]{view['id']}[/bold]  ({view['category']}, {view['difficulty']})")
    console.print(f"[dim]answer_type: {view['answer_type']}[/dim]\n")
    console.print(view["question"])


@task_app.command("submit")
def submit_cmd(
    task_id: str = typer.Argument(...),
    answer: str = typer.Option(..., "--answer", "-a", help="Your final answer."),
    reveal: bool = typer.Option(False, "--reveal", help="Show the gold answer on failure."),
) -> None:
    """Grade a final answer for one task. Exit code 0 = correct, 1 = wrong."""
    try:
        task = runner.get_task(task_id)
    except KeyError as exc:
        console.print(f"[red]{exc.args[0]}[/red]")
        raise typer.Exit(code=2) from exc
    result = grade(task, answer)
    runner.log_submission(task["id"], answer, result)
    if result.correct:
        console.print(f"[bold green]✓ CORRECT[/bold green] — {task['id']}")
        raise typer.Exit(code=0)
    console.print(f"[bold red]✗ INCORRECT[/bold red] — {task['id']}")
    if result.note:
        console.print(f"[dim]{result.note}[/dim]")
    if reveal:
        console.print(f"expected: {result.expected}")
    raise typer.Exit(code=1)


@task_app.command("grade-file")
def grade_file_cmd(
    answers_file: Path = typer.Argument(
        ..., exists=True, readable=True,
        help='JSON mapping {"HOSP-001": "10", …} or JSONL of {"id": …, "answer": …}.',
    ),
) -> None:
    """Batch-grade a whole answer file and print a scoreboard."""
    raw = answers_file.read_text()
    answers: dict[str, str] = {}
    try:
        answers = {str(k): str(v) for k, v in json.loads(raw).items()}
    except json.JSONDecodeError:
        for line in raw.splitlines():
            line = line.strip()
            if not line:
                continue
            record = json.loads(line)
            answers[str(record["id"])] = str(record["answer"])

    tasks = {t["id"]: t for t in runner.load_tasks()}
    table = Table(title="Grade report")
    for header in ("task", "verdict", "submitted"):
        table.add_column(header)
    correct = 0
    for task_id, submitted in answers.items():
        task = tasks.get(task_id)
        if task is None:
            table.add_row(task_id, "[yellow]unknown id[/yellow]", submitted)
            continue
        result = grade(task, submitted)
        correct += int(result.correct)
        verdict = "[green]✓[/green]" if result.correct else "[red]✗[/red]"
        table.add_row(task_id, verdict, submitted)
    console.print(table)
    graded = sum(1 for tid in answers if tid in tasks)
    pct = (100.0 * correct / graded) if graded else 0.0
    console.print(f"[bold]score: {correct}/{graded} ({pct:.0f}%)[/bold]")


@task_app.command("check")
def check_cmd(
    task_id: str | None = typer.Argument(None, help="Verify one task; omit for all."),
) -> None:
    """MAINTAINER: re-derive every gold answer from the live DBs and compare."""
    tasks = [runner.get_task(task_id)] if task_id else runner.load_tasks()
    table = Table(title="Gold verification")
    for header in ("task", "stored gold", "derived", "match"):
        table.add_column(header, overflow="fold")
    failures = 0
    for task in tasks:
        try:
            derived, result = runner.check_task(task)
        except Exception as exc:  # noqa: BLE001 — report and continue the sweep
            failures += 1
            table.add_row(task["id"], str(task.get("gold_answer")), f"ERROR: {exc}", "[red]✗[/red]")
            continue
        failures += 0 if result.correct else 1
        mark = "[green]✓[/green]" if result.correct else "[red]✗[/red]"
        table.add_row(task["id"], str(task.get("gold_answer")), derived, mark)
    console.print(table)
    if failures:
        console.print(f"[bold red]{failures} task(s) failed verification.[/bold red]")
        raise typer.Exit(code=1)
    console.print(f"[bold green]All {len(tasks)} gold answers verified against live data.[/bold green]")


@task_app.command("export-questions")
def export_questions_cmd(
    out: Path = typer.Option(None, "--out", "-o", help="Write to file instead of stdout."),
) -> None:
    """Emit the agent-safe questions-only JSONL (no golds, no reference SQL)."""
    lines = [json.dumps(runner.public_view(t)) for t in runner.load_tasks()]
    payload = "\n".join(lines) + "\n"
    if out:
        out.write_text(payload)
        console.print(f"[green]✓ Wrote {len(lines)} questions to {out}[/green]")
    else:
        typer.echo(payload, nl=False)
