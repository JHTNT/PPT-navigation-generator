from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator, List, Optional, Sequence, Tuple


@dataclass(frozen=True)
class OutlineItem:
    """Represents a single entry in the markdown outline."""

    title: str
    children: Tuple["OutlineItem", ...] = field(default_factory=tuple)

    @property
    def has_children(self) -> bool:
        return bool(self.children)


@dataclass(frozen=True)
class SlidePlanEntry:
    """Describes the section/subsection pairing that becomes one PPT slide."""

    section: OutlineItem
    child: Optional[OutlineItem]


@dataclass(frozen=True)
class Outline:
    """Full outline along with utility helpers for downstream builders."""

    sections: Tuple[OutlineItem, ...]

    @classmethod
    def from_text(cls, text: str) -> "Outline":
        return cls(sections=_parse_markdown_lines(text.splitlines()))

    @classmethod
    def from_file(cls, path: Path) -> "Outline":
        return cls.from_text(path.read_text(encoding="utf-8"))

    def iter_slide_plan(self) -> Iterator[SlidePlanEntry]:
        for section in self.sections:
            if section.children:
                for child in section.children:
                    yield SlidePlanEntry(section=section, child=child)
            else:
                yield SlidePlanEntry(section=section, child=None)


def _parse_markdown_lines(lines: Sequence[str]) -> Tuple[OutlineItem, ...]:
    sections: List[_MutableNode] = []
    stack: List[_MutableNode] = []

    for line_number, raw_line in enumerate(lines, start=1):
        line = raw_line.rstrip()
        if not line:
            continue
        stripped = line.lstrip()
        if not stripped.startswith("- "):
            continue

        indent = len(line) - len(stripped)
        if indent % 2 != 0:
            raise ValueError(f"Line {line_number}: indentation must use multiples of two spaces.")
        level = indent // 2 + 1
        if level > 2:
            raise ValueError("Only two heading levels are supported.")

        title = stripped[2:].strip()
        if not title:
            raise ValueError(f"Line {line_number}: bullet items must have a title.")

        node = _MutableNode(title=title, level=level)

        while len(stack) >= level:
            stack.pop()

        if level == 1:
            sections.append(node)
        else:
            if not stack:
                raise ValueError(
                    f"Line {line_number}: found a second-level item without a parent section."
                )
            stack[-1].children.append(node)
        stack.append(node)

    return tuple(_freeze(node) for node in sections)


@dataclass
class _MutableNode:
    title: str
    level: int
    children: List["_MutableNode"] = field(default_factory=list)


def _freeze(node: _MutableNode) -> OutlineItem:
    return OutlineItem(
        title=node.title,
        children=tuple(_freeze(child) for child in node.children),
    )
