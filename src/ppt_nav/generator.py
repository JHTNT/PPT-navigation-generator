from __future__ import annotations

"""High-level helpers that tie parsing and presentation building together."""

from pathlib import Path
from typing import Optional

from ppt_nav.outline import Outline
from ppt_nav.ppt_builder import PresentationBuilder


def generate_from_markdown(
    input_path: Path,
    output_path: Optional[Path] = None,
    font_size: Optional[float] = None,
    template_path: Optional[Path] = None,
) -> Path:
    """Parse ``input_path`` and build a PPTX deck at ``output_path``.

    ``font_size`` is the base font size in points used for navigation tabs and
    the main content placeholder.

    ``template_path`` is an optional PPTX template to load. When omitted, the
    bundled ``template/template_16-9.pptx`` is used when present.
    """

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    outline = Outline.from_file(input_path)
    builder = PresentationBuilder(font_size=font_size)
    destination = output_path or input_path.with_suffix(".pptx")

    resolved_template: Path
    if template_path is not None:
        if not template_path.exists():
            raise FileNotFoundError(f"Template file not found: {template_path}")
        resolved_template = template_path
    else:
        # Default to the repo-provided 16:9 template.
        default_template = (
            Path(__file__).resolve().parents[2] / "template" / "template_16-9.pptx"
        )
        if not default_template.exists():
            raise FileNotFoundError(
                "Default template not found: template/template_16-9.pptx. "
                "Provide one via --template."
            )
        resolved_template = default_template

    builder.build(outline, destination, template_path=resolved_template)
    return destination
