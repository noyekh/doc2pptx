"""Command-line interface for **doc2pptx**

Fonctionnalités :
* argument positionnel JSON, option `--template` obligatoire,
* `--output` facultatif (défaut : <INPUT>.pptx),
* logs Rich (`--log-level` OU `--verbose/-v`),
* barre de progression Rich,
* sous-commandes : generate / edit (stub) / prompt (stub) / version.
"""

from __future__ import annotations

import json
import logging
from importlib.metadata import PackageNotFoundError, version as pkg_version
from pathlib import Path
from typing import Optional, Union

import typer
from rich.console import Console
from rich.logging import RichHandler
from rich.progress import Progress

from doc2pptx.core.models import Presentation
from doc2pptx.ingest.json_loader import load_presentation
from doc2pptx.ppt.builder import PPTBuilder

console = Console()
app = typer.Typer(add_completion=False, no_args_is_help=True)


# ──────────────────────────────── utils ────────────────────────────────
def _configure_logging(level: Union[str, bool] = "INFO") -> None:
    """Configure le root logger.

    - str  : "DEBUG" / "INFO" / …
    - bool : True → DEBUG, False → INFO  (compat tests)"""
    if isinstance(level, bool):
        level_str = "DEBUG" if level else "INFO"
    else:
        level_str = level.upper()

    logging.basicConfig(
        level=getattr(logging, level_str, logging.INFO),
        format="%(message)s",
        handlers=[RichHandler(rich_tracebacks=True, markup=True)],
    )
    logging.getLogger().setLevel(getattr(logging, level_str, logging.INFO))

import builtins  # après logging
builtins._configure_logging = _configure_logging


def _default_output_for(input_path: Path) -> Path:
    return (Path.cwd() / input_path.with_suffix(".pptx").name).resolve()


def _build_presentation(
    json_path: Path, template_path: Path, output_path: Path
) -> Path:
    """Parse JSON → build PPTX → return generated path (ou lève)."""
    presentation: Presentation = load_presentation(json_path)
    builder = PPTBuilder(template_path=template_path)
    with Progress(transient=True) as progress:
        task = progress.add_task("[green]Construction…", total=None)
        generated = builder.build(presentation, output_path)
        progress.update(task, completed=1)
    return generated


# ─────────────────────────────── generate ──────────────────────────────
@app.command("generate")
def generate(  # noqa: D401 – imperative form for CLI
    input_json: Path = typer.Argument(
        ...,
        exists=True,
        readable=True,
        resolve_path=True,
        metavar="INPUT_JSON",
        help="Fichier JSON structuré décrivant la présentation.",
    ),
    template: Path = typer.Option(
        ..., "--template", "-t", exists=True, readable=True, resolve_path=True
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output",
        "-o",
        writable=True,
        resolve_path=True,
        metavar="OUTPUT_PPTX",
        help="Fichier de sortie (.pptx). Défaut : <INPUT>.pptx",
    ),
    log_level: str = typer.Option(
        "INFO",
        "--log-level",
        "-l",
        help="Niveau de log (DEBUG, INFO, WARNING, ERROR).",
        show_default=False,
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose",
        "-v",
        help="Alias pratique pour --log-level DEBUG.",
    ),
) -> None:
    """Générer un deck PowerPoint à partir d’un JSON structuré."""
    effective_level = "DEBUG" if verbose else log_level
    _configure_logging(effective_level)

    output = output or _default_output_for(input_json)
    output.parent.mkdir(parents=True, exist_ok=True)

    console.rule("[bold cyan]doc2pptx • generate")
    console.print(f"Template   : [cyan]{template}[/cyan]")
    console.print(f"Entrée     : [magenta]{input_json}[/magenta]")
    console.print(f"Destination: [green]{output}[/green]")

    try:
        generated_path = _build_presentation(input_json, template, output)
        console.print(f"[bold green]Deck généré :[/] {generated_path}")
    except json.JSONDecodeError as exc:
        console.print(f"[bold red]Erreur de parsing JSON :[/] {exc}")
        raise typer.Exit(code=1) from exc
    except Exception as exc:  # noqa: BLE001 – capte Pydantic ou builder
        console.print(f"[bold red]Échec génération :[/] {exc}")
        if effective_level.upper() == "DEBUG":
            console.print_exception()
        raise typer.Exit(code=2) from exc


# ──────────────────────────── stubs Jour 7 ─────────────────────────────
@app.command()
def edit(
    presentation: Path = typer.Option(
        ...,
        "--presentation",
        "-p",
        exists=False,
        resolve_path=True,
        help="Fichier .pptx à modifier.",
    ),
    command: str = typer.Option(
        ..., "--command", "-c", help='Commande d’édition (ex. "move slide 2 to 5").'
    ),
) -> None:
    """Modifier une présentation — sera disponible Jour 7."""
    console.print(
        "[yellow]La commande [bold]edit[/] n’est pas encore disponible (roadmap Jour 7).[/yellow]"
    )
    raise typer.Exit(code=1)


@app.command()
def prompt(
    presentation: Path = typer.Option(
        ...,
        "--presentation",
        "-p",
        exists=False,
        resolve_path=True,
        help="Fichier .pptx à modifier.",
    ),
    nl: str = typer.Option(..., "--nl", help="Commande en langage naturel."),
) -> None:
    """Appliquer des instructions NL — sera disponible Jour 7."""
    console.print(
        "[yellow]La commande [bold]prompt[/] n’est pas encore disponible (roadmap Jour 7).[/yellow]"
    )
    raise typer.Exit(code=1)


# ─────────────────────────────── version ───────────────────────────────
@app.command()
def version() -> None:  # noqa: D401
    """Afficher la version installée de *doc2pptx*."""
    try:
        console.print(f"doc2pptx : [bold green]{pkg_version('doc2pptx')}[/]")
    except PackageNotFoundError:
        console.print("[yellow]Package non installé (mode editable ou tests).[/]")


# ───────────────────────────── entry-point ─────────────────────────────
def _entrypoint() -> None:  # pragma: no cover
    app()


if __name__ == "__main__":  # pragma: no cover
    _entrypoint()