from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass
class Run:
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False


_TOKEN = re.compile(
    r"(\*\*([^*]+)\*\*)"         # **bold**
    r"|(\*([^*]+)\*)"            # *italic*
    r"|(`([^`]+)`)"              # `code`
    r"|(\[([^\]]+)\]\(#ref(\d+)\))",  # [text](#refN) -> [N]
)


def parse_runs(text: str) -> list[Run]:
    """Parse inline Markdown into a list of styled runs.

    Supports: **bold**, *italic*, `code`, [anything](#refN) -> [N].
    Unknown markup is left as literal text.
    """
    runs: list[Run] = []
    pos = 0
    for m in _TOKEN.finditer(text):
        if m.start() > pos:
            runs.append(Run(text[pos:m.start()]))
        if m.group(1):
            runs.append(Run(m.group(2), bold=True))
        elif m.group(3):
            runs.append(Run(m.group(4), italic=True))
        elif m.group(5):
            runs.append(Run(m.group(6), code=True))
        elif m.group(7):
            runs.append(Run(f"[{m.group(9)}]"))
        pos = m.end()
    if pos < len(text):
        runs.append(Run(text[pos:]))
    return runs


def split_paragraphs(body: str) -> list[str]:
    """Split body text on blank lines, then normalise internal whitespace.

    Hard-wrapped prose in the Markdown source would otherwise render as a run
    of <w:br/> soft returns in the docx — which under IEEE's justified body
    style produces visibly stretched lines. Collapse any internal whitespace
    run (including single newlines) to a single space so each paragraph flows
    as one continuous line.
    """
    chunks = re.split(r"\n\s*\n", body.strip())
    return [re.sub(r"\s+", " ", c).strip() for c in chunks if c.strip()]


_BULLET_LINE = re.compile(r"^\s*[-*]\s+(.*)$")
_EQUATION_BLOCK = re.compile(r"^\s*\$\$(.+?)\$\$\s*$", re.DOTALL)


def parse_body_blocks(body: str) -> list[tuple[str, str]]:
    """Parse body text into a sequence of ('para'|'bullet'|'equation', text)
    blocks.

    Blocks are separated by blank lines. A chunk is classified as:
      * 'equation' if it is wrapped in `$$…$$`.
      * 'bullet' if every non-empty line starts with `-` or `*` (one bullet
        entry per line).
      * 'para' otherwise, with internal whitespace collapsed to single spaces.
    """
    chunks = re.split(r"\n\s*\n", body.strip())
    blocks: list[tuple[str, str]] = []
    for chunk in chunks:
        eq = _EQUATION_BLOCK.match(chunk)
        if eq:
            blocks.append(("equation", re.sub(r"\s+", " ", eq.group(1)).strip()))
            continue
        lines = [l for l in chunk.split("\n") if l.strip()]
        if not lines:
            continue
        if all(_BULLET_LINE.match(l) for l in lines):
            for line in lines:
                text = _BULLET_LINE.match(line).group(1).strip()
                blocks.append(("bullet", re.sub(r"\s+", " ", text)))
        else:
            blocks.append(("para", re.sub(r"\s+", " ", chunk).strip()))
    return blocks
