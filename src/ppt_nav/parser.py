from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List


@dataclass
class OutlineNode:
    title: str
    level: int
    children: List["OutlineNode"] = field(default_factory=list)


@dataclass
class OutlineDocument:
    sections: List[OutlineNode]


class OutlineParser:
    def parse_file(self, path: Path) -> OutlineDocument:
        text = path.read_text(encoding="utf-8")
        sections = self._parse_body(text)
        return OutlineDocument(sections=sections)

    def _parse_body(self, body: str) -> List[OutlineNode]:
        sections: List[OutlineNode] = []
        stack: List[OutlineNode] = []
        for raw_line in body.splitlines():
            line = raw_line.rstrip()
            if not line:
                continue
            stripped = line.lstrip()
            if not stripped.startswith("- "):
                continue
            indent = len(line) - len(stripped)
            level = indent // 2 + 1
            if level > 2:
                raise ValueError("Only two levels of headings are supported")
            title = stripped[2:].strip()
            node = OutlineNode(title=title, level=level)
            while len(stack) >= level:
                stack.pop()
            if level == 1:
                sections.append(node)
            else:
                if not stack:
                    raise ValueError("Sub-item found without a parent section")
                parent = stack[-1]
                parent.children.append(node)
            stack.append(node)
        return sections
