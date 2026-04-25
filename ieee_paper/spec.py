from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass
class Author:
    name: str
    dept: str = ""
    org: str = ""
    city: str = ""
    country: str = ""
    email: str = ""


@dataclass
class Figure:
    id: str = ""
    path: str = ""  # empty → placeholder caption only, no image
    caption: str = ""


@dataclass
class Cell:
    text: str = ""
    rowspan: int = 1
    colspan: int = 1
    style: str = ""  # optional override of row-derived style


@dataclass
class Table:
    caption: str
    rows: list[list[Cell]] = field(default_factory=list)
    has_subhead_row: bool = False
    footnote: str = ""


@dataclass
class Section:
    heading: str
    body: str = ""
    subsections: list[Section] = field(default_factory=list)
    numbered: bool = True  # False → un-numbered Heading 5 (Acknowledgment, etc.)
    figures: list[Figure] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)


@dataclass
class PaperSpec:
    title: str
    authors: list[Author]
    abstract: str
    keywords: list[str]
    sections: list[Section]
    figures: list[Figure] = field(default_factory=list)
    references: list[str] = field(default_factory=list)
    references_numbered: bool = False
    references_body: str = ""
    sponsors: str = ""
    trailing_text: str = ""
    base_dir: Path = field(default_factory=Path)

    @classmethod
    def from_yaml(cls, path: str | Path) -> PaperSpec:
        path = Path(path)
        with path.open() as f:
            raw = yaml.safe_load(f)
        return cls.from_dict(raw, base_dir=path.parent)

    @classmethod
    def from_dict(cls, d: dict[str, Any], base_dir: Path = Path(".")) -> PaperSpec:
        authors = [Author(**a) for a in d.get("authors", [])]
        sections = [_parse_section(s, base_dir) for s in d.get("sections", [])]
        figures = [Figure(**f) for f in d.get("figures", [])]
        return cls(
            title=d["title"],
            authors=authors,
            abstract=_resolve_text(d.get("abstract", ""), base_dir),
            keywords=list(d.get("keywords", [])),
            sections=sections,
            figures=figures,
            references=list(d.get("references", [])),
            references_numbered=bool(d.get("references_numbered", False)),
            references_body=_resolve_text(d.get("references_body", ""), base_dir),
            sponsors=d.get("sponsors", ""),
            trailing_text=_resolve_text(d.get("trailing_text", ""), base_dir),
            base_dir=base_dir,
        )


def _parse_section(d: dict[str, Any], base_dir: Path) -> Section:
    return Section(
        heading=d["heading"],
        body=_resolve_text(d.get("body", ""), base_dir),
        subsections=[_parse_section(s, base_dir) for s in d.get("subsections", [])],
        numbered=bool(d.get("numbered", True)),
        figures=[Figure(**f) for f in d.get("figures", [])],
        tables=[_parse_table(t) for t in d.get("tables", [])],
    )


def _parse_table(d: dict[str, Any]) -> Table:
    return Table(
        caption=d.get("caption", ""),
        rows=[[_parse_cell(c) for c in r] for r in d.get("rows", [])],
        has_subhead_row=bool(d.get("has_subhead_row", False)),
        footnote=d.get("footnote", ""),
    )


def _parse_cell(c: Any) -> Cell:
    if isinstance(c, Cell):
        return c
    if isinstance(c, dict):
        return Cell(
            text=c.get("text", ""),
            rowspan=int(c.get("rowspan", 1)),
            colspan=int(c.get("colspan", 1)),
            style=c.get("style", ""),
        )
    return Cell(text=str(c))


def _resolve_text(value: str, base_dir: Path) -> str:
    """If value looks like a file path to a .md/.txt file that exists, load it; else treat as literal text."""
    if not value:
        return ""
    stripped = value.strip()
    if len(stripped) < 260 and "\n" not in stripped and stripped.endswith((".md", ".txt")):
        candidate = (base_dir / stripped).resolve()
        if candidate.is_file():
            return candidate.read_text()
    return value
