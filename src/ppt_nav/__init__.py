"""Public API for the PPT navigation generator package."""

from ppt_nav.generator import generate_from_markdown
from ppt_nav.outline import Outline, OutlineItem

__all__ = [
	"generate_from_markdown",
	"Outline",
	"OutlineItem",
]
