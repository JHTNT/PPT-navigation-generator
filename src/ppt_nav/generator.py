from __future__ import annotations

"""High-level helpers that tie parsing and presentation building together."""

from pathlib import Path
from typing import Optional

from ppt_nav.outline import Outline
from ppt_nav.ppt_builder import PresentationBuilder


def generate_from_markdown(input_path: Path, output_path: Optional[Path] = None) -> Path:
    """Parse ``input_path`` and build a PPTX deck at ``output_path``."""

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    outline = Outline.from_file(input_path)
    builder = PresentationBuilder()
    destination = output_path or input_path.with_suffix(".pptx")
    builder.build(outline, destination)
    return destination
