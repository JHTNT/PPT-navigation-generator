from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from ppt_nav.generator import generate_from_markdown


def run(argv: Sequence[str] | None = None) -> int:
    """Entry point used by both ``python -m ppt_nav`` and ``ppt-nav``."""

    parsed_args = _build_parser().parse_args(
        argv if argv is not None else sys.argv[1:])
    return _handle_build(parsed_args)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ppt-nav",
        description="Generate PPT decks with a two-layer navigation header.",
    )
    parser.add_argument("input", type=Path, help="Markdown outline file.")
    parser.add_argument(
        "output",
        type=Path,
        nargs="?",
        help="Optional PPTX path (defaults to <input>.pptx).",
    )
    parser.add_argument(
        "--font-size",
        type=float,
        default=28.0,
        help="Base font size in points for navigation and content (default: 28).",
    )
    return parser


def _handle_build(args: argparse.Namespace) -> int:
    input_path: Path = args.input
    output_path: str | None = args.output
    font_size: float = args.font_size

    try:
        if font_size <= 0:
            raise ValueError("Font size must be positive")
        destination = generate_from_markdown(input_path, output_path, font_size=font_size)
    except FileNotFoundError as exc:
        print(str(exc))
        return 1
    except ValueError as exc:
        print(f"Outline error: {exc}")
        return 1

    print(f"Presentation generated at {destination}")
    return 0
