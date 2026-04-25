"""Render a PaperSpec to a standalone .tex file using IEEE's IEEEtran.cls.

The output is a complete LaTeX source ready for `pdflatex`. It relies on
`IEEEtran.cls` being on the user's LaTeX path; the class ships with every
standard TeX distribution (TeX Live, MacTeX, MiKTeX) and with IEEE's own
template bundle.

This generator is deliberately thin: LaTeX + IEEEtran handle all the layout
rules (column flow, section numbering, author grid, equation/table/figure
placement, reference list formatting) that our .docx generator has to
reconstruct by hand in OOXML.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

from .markdown_parser import parse_body_blocks, parse_runs
from .spec import Author, Cell, Figure, PaperSpec, Section, Table


# --- Entry point ---------------------------------------------------------


def build_paper_latex(spec: PaperSpec, output_path: str | Path) -> None:
    """Render `spec` to a .tex file at `output_path`."""
    output_path = Path(output_path)
    source = _render(spec)
    output_path.write_text(source)


# --- Top-level renderer --------------------------------------------------


def _render(spec: PaperSpec) -> str:
    parts: list[str] = []
    parts.append(_preamble(spec))
    parts.append(r"\begin{document}")
    parts.append(_title_block(spec))
    parts.append(r"\maketitle")
    parts.append("")
    parts.append(_abstract_block(spec))
    parts.append(_keywords_block(spec))

    for section in spec.sections:
        parts.append(_render_section(section, depth=0))

    if spec.references:
        parts.append(_references_block(spec.references, spec.references_body))

    if spec.trailing_text.strip():
        parts.append(_render_body(spec.trailing_text))

    parts.append(r"\end{document}")
    return "\n".join(parts) + "\n"


# --- Preamble ------------------------------------------------------------


def _preamble(spec: PaperSpec) -> str:
    # Match IEEE's recommended preamble from the official template, plus
    # `multirow` for tables that use rowspan.
    lines = [
        r"\documentclass[conference]{IEEEtran}",
        r"\IEEEoverridecommandlockouts",
        r"\usepackage{cite}",
        r"\usepackage{amsmath,amssymb,amsfonts}",
        r"\usepackage{graphicx}",
        r"\usepackage{textcomp}",
        r"\usepackage{xcolor}",
        r"\usepackage{url}",
        r"\usepackage{multirow}",
    ]
    return "\n".join(lines)


# --- Title + authors -----------------------------------------------------


def _title_block(spec: PaperSpec) -> str:
    title_content = _render_inline(spec.title)
    if spec.sponsors:
        title_content += r"\thanks{" + _escape(spec.sponsors) + "}"
    title_block = r"\title{" + title_content + "}"
    author_block = _authors_block(spec.authors)
    return title_block + "\n\n" + author_block


_LEADING_ORDINAL = re.compile(r"^\s*\d+(?:st|nd|rd|th)\s+", re.IGNORECASE)


def _authors_block(authors: list[Author]) -> str:
    if not authors:
        return r"\author{}"
    blocks: list[str] = []
    for i, a in enumerate(authors, start=1):
        # If the user wrote "1st Given Name Surname" in the YAML (matching the
        # docx template's literal placeholder text), strip the leading ordinal
        # so we don't duplicate it with our typographic `\textsuperscript`.
        bare_name = _LEADING_ORDINAL.sub("", a.name)
        # LaTeX expects single-line dept/org; collapse any embedded newlines
        # that the docx-era YAML uses for the two-line "(of Affiliation)"
        # continuation.
        dept = _single_line(a.dept)
        org = _single_line(a.org)

        lines_after_name: list[str] = []
        if dept.strip():
            lines_after_name.append(r"\textit{" + _escape(dept) + "}")
        if org.strip():
            lines_after_name.append(r"\textit{" + _escape(org) + "}")
        cc = _city_country(a)
        if cc:
            lines_after_name.append(_escape(cc))
        if a.email.strip():
            lines_after_name.append(_escape(a.email))

        block = (
            r"\IEEEauthorblockN{" + str(i)
            + r"\textsuperscript{" + _superscript_suffix(i) + "} "
            + _escape(bare_name) + "}\n"
            + r"\IEEEauthorblockA{"
            + " \\\\\n".join(lines_after_name)
            + "}"
        )
        blocks.append(block)
    joined = "\n\\and\n".join(blocks)
    return r"\author{" + "\n" + joined + "\n}"


def _single_line(s: str) -> str:
    """Collapse any whitespace (including newlines) into single spaces. Used
    for author fields that must fit on one line in the LaTeX template."""
    return re.sub(r"\s+", " ", s).strip()


def _superscript_suffix(n: int) -> str:
    if 10 <= n % 100 <= 20:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")


def _city_country(a: Author) -> str:
    parts = [a.city.strip(), a.country.strip()]
    return ", ".join(p for p in parts if p)


# --- Abstract / keywords --------------------------------------------------


def _abstract_block(spec: PaperSpec) -> str:
    if not spec.abstract.strip():
        return ""
    body = _escape(re.sub(r"\s+", " ", spec.abstract).strip())
    return r"\begin{abstract}" + "\n" + body + "\n" + r"\end{abstract}"


def _keywords_block(spec: PaperSpec) -> str:
    if not spec.keywords:
        return ""
    joined = ", ".join(_escape(k) for k in spec.keywords if k.strip())
    return r"\begin{IEEEkeywords}" + "\n" + joined + "\n" + r"\end{IEEEkeywords}"


# --- Sections ------------------------------------------------------------


_SECTION_CMDS = [
    r"\section",       # depth 0 → Heading 1 (Roman numeral)
    r"\subsection",    # depth 1 → Heading 2 (A., B.)
    r"\subsubsection", # depth 2 → Heading 3 (1), 2))
    r"\paragraph",     # depth 3 → Heading 4 (a), b))
    r"\subparagraph",  # depth 4 → Heading 5 (un-numbered)
]


def _render_section(section: Section, depth: int) -> str:
    cmd = _SECTION_CMDS[min(depth, len(_SECTION_CMDS) - 1)]
    starred = "" if section.numbered else "*"
    parts: list[str] = []
    parts.append(cmd + starred + "{" + _render_inline(section.heading) + "}")
    parts.append(_render_body(section.body))
    for table in section.tables:
        parts.append(_render_table(table))
    for fig in section.figures:
        parts.append(_render_figure(fig))
    for sub in section.subsections:
        parts.append(_render_section(sub, depth + 1))
    return "\n".join(p for p in parts if p.strip())


def _render_body(body: str) -> str:
    if not body.strip():
        return ""
    out: list[str] = []
    bullet_buffer: list[str] = []

    def flush_bullets() -> None:
        if not bullet_buffer:
            return
        out.append(r"\begin{itemize}")
        for b in bullet_buffer:
            out.append(r"\item " + b)
        out.append(r"\end{itemize}")
        bullet_buffer.clear()

    for kind, text in parse_body_blocks(body):
        if kind == "bullet":
            bullet_buffer.append(_render_inline(text))
            continue
        flush_bullets()
        if kind == "equation":
            out.append(r"\begin{equation}")
            # Strip our own `*italic*` markdown from the equation — LaTeX
            # already italicises Roman letters inside math mode.
            math = re.sub(r"\*([^*]+)\*", r"\1", text)
            out.append(_translate_math_unicode(math))
            out.append(r"\end{equation}")
        else:  # 'para'
            out.append(_render_inline(text))
            out.append("")  # blank line between paragraphs
    flush_bullets()
    return "\n".join(out).rstrip()


def _render_inline(text: str) -> str:
    """Convert inline markdown runs to LaTeX, escaping special characters."""
    pieces: list[str] = []
    for run in parse_runs(text):
        escaped = _escape(run.text)
        if run.code:
            pieces.append(r"\texttt{" + escaped + "}")
        elif run.bold and run.italic:
            pieces.append(r"\textbf{\textit{" + escaped + "}}")
        elif run.bold:
            pieces.append(r"\textbf{" + escaped + "}")
        elif run.italic:
            pieces.append(r"\textit{" + escaped + "}")
        else:
            pieces.append(escaped)
    return "".join(pieces)


# --- Figures -------------------------------------------------------------


def _render_figure(fig: Figure) -> str:
    lines = [r"\begin{figure}[htbp]", r"\centerline{"]
    if fig.path:
        lines.append(r"\includegraphics[width=\columnwidth]{" + fig.path + "}")
    else:
        # Placeholder rule so the caption still reads sensibly without an image
        lines.append(r"\fbox{\parbox{0.8\columnwidth}{\centering \vspace{2em} [figure placeholder] \vspace{2em}}}")
    lines.append(r"}")
    lines.append(r"\caption{" + _render_inline(fig.caption) + "}")
    if fig.id:
        lines.append(r"\label{" + _escape(fig.id) + "}")
    lines.append(r"\end{figure}")
    return "\n".join(lines)


# --- Tables --------------------------------------------------------------


def _render_table(table: Table) -> str:
    if not table.rows:
        return ""
    placements = _compute_placements(table.rows)
    nrows = len(table.rows)
    ncols = max(c + cs for (_, c, _, cs, _) in placements)

    # Build per-start-cell rendering tokens plus two coverage sets so we can
    # decide at each (r, c) position: emit the token, skip silently (covered
    # horizontally by a same-row colspan — `\multicolumn` already filled the
    # LaTeX columns), or emit an empty cell (covered vertically by a prior
    # row's rowspan — we still need the column separator for alignment).
    start_cells: dict[tuple[int, int], tuple[int, int, str]] = {}
    col_covered: set[tuple[int, int]] = set()
    row_covered: set[tuple[int, int]] = set()

    for r, c, rowspan, colspan, cell in placements:
        content = _render_inline(cell.text)
        if rowspan > 1 and colspan > 1:
            token = (
                r"\multicolumn{" + str(colspan) + r"}{|c|}{"
                + r"\multirow{" + str(rowspan) + r"}{*}{" + content + "}}"
            )
        elif colspan > 1:
            token = r"\multicolumn{" + str(colspan) + r"}{|c|}{" + content + "}"
        elif rowspan > 1:
            token = r"\multirow{" + str(rowspan) + r"}{*}{" + content + "}"
        else:
            token = content
        start_cells[(r, c)] = (rowspan, colspan, token)
        for dr in range(rowspan):
            for dc in range(colspan):
                if dr == 0 and dc == 0:
                    continue
                if dr == 0:
                    col_covered.add((r, c + dc))
                else:
                    row_covered.add((r + dr, c + dc))

    col_spec = "|" + "|".join(["c"] * ncols) + "|"

    lines: list[str] = [
        r"\begin{table}[htbp]",
        r"\caption{" + _render_inline(table.caption) + "}",
        r"\begin{center}",
        r"\begin{tabular}{" + col_spec + "}",
        r"\hline",
    ]

    for r in range(nrows):
        cells: list[str] = []
        c = 0
        while c < ncols:
            if (r, c) in start_cells:
                rowspan, colspan, token = start_cells[(r, c)]
                cells.append(token)
                c += colspan
            elif (r, c) in col_covered:
                c += 1  # colspan already consumed this column in LaTeX
            else:
                # Either a row-covered (below a rowspan) or plain empty slot;
                # either way we need an empty logical cell for alignment.
                cells.append("")
                c += 1
        lines.append(" & ".join(cells) + r" \\")
        lines.append(r"\hline")

    lines.append(r"\end{tabular}")
    if table.footnote:
        lines.append(r"\\[2pt] \footnotesize " + _render_inline(table.footnote))
    lines.append(r"\end{center}")
    lines.append(r"\end{table}")
    return "\n".join(lines)


def _compute_placements(
    rows: list[list[Cell]],
) -> list[tuple[int, int, int, int, Cell]]:
    """Same placement algorithm as the docx builder — walk rows left→right,
    skipping positions already covered by prior rowspans."""
    reserved: set[tuple[int, int]] = set()
    placements: list[tuple[int, int, int, int, Cell]] = []
    for r, row in enumerate(rows):
        c = 0
        for cell in row:
            while (r, c) in reserved:
                c += 1
            placements.append((r, c, cell.rowspan, cell.colspan, cell))
            for dr in range(cell.rowspan):
                for dc in range(cell.colspan):
                    if dr == 0 and dc == 0:
                        continue
                    reserved.add((r + dr, c + dc))
            c += cell.colspan
    return placements


# --- References ----------------------------------------------------------


def _references_block(refs: list[str], body: str = "") -> str:
    lines: list[str] = []
    if body.strip():
        lines.append(_render_body(body))
    lines.append(r"\begin{thebibliography}{" + str(max(len(refs), 99)) + "}")
    for i, ref in enumerate(refs, start=1):
        cleaned = re.sub(r"^\s*\[\d+\]\s*", "", ref)
        lines.append(r"\bibitem{b" + str(i) + "} " + _render_inline(cleaned))
    lines.append(r"\end{thebibliography}")
    return "\n".join(lines)


# --- Escaping ------------------------------------------------------------


# Map of LaTeX-special characters that must be escaped in ordinary text.
_LATEX_ESCAPES = {
    "\\": r"\textbackslash{}",
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "_": r"\_",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}

# Unicode characters that pdflatex (with the default OT1 encoding) can't
# render directly. We wrap them in inline math or use textcomp equivalents so
# the author can type α/μ/₀ in the YAML and still compile cleanly. Keep this
# list narrow — only characters the template demo actually uses.
_UNICODE_TEXT_MAP: dict[str, str] = {
    "α": r"$\alpha$",
    "β": r"$\beta$",
    "γ": r"$\gamma$",
    "δ": r"$\delta$",
    "μ": r"$\mu$",
    "π": r"$\pi$",
    "σ": r"$\sigma$",
    "χ": r"$\chi$",
    "ω": r"$\omega$",
    "₀": r"$_0$",
    "₁": r"$_1$",
    "₂": r"$_2$",
    "₃": r"$_3$",
    "×": r"$\times$",
    "÷": r"$\div$",
    "±": r"$\pm$",
    "≈": r"$\approx$",
    "≤": r"$\leq$",
    "≥": r"$\geq$",
    "≠": r"$\neq$",
    "∞": r"$\infty$",
    "→": r"$\rightarrow$",
    "—": r"---",
    "–": r"--",
    "“": "``",
    "”": "''",
    "‘": "`",
    "’": "'",
}

_UNICODE_MATH_MAP: dict[str, str] = {
    "α": r"\alpha ",
    "β": r"\beta ",
    "γ": r"\gamma ",
    "δ": r"\delta ",
    "μ": r"\mu ",
    "π": r"\pi ",
    "σ": r"\sigma ",
    "χ": r"\chi ",
    "ω": r"\omega ",
    "₀": r"_{0}",
    "₁": r"_{1}",
    "₂": r"_{2}",
    "₃": r"_{3}",
    "×": r"\times ",
    "÷": r"\div ",
    "±": r"\pm ",
    "≈": r"\approx ",
    "≤": r"\leq ",
    "≥": r"\geq ",
    "≠": r"\neq ",
    "∞": r"\infty ",
    "→": r"\rightarrow ",
}


def _translate_text_unicode(s: str) -> str:
    return "".join(_UNICODE_TEXT_MAP.get(ch, ch) for ch in s)


def _translate_math_unicode(s: str) -> str:
    return "".join(_UNICODE_MATH_MAP.get(ch, ch) for ch in s)


def _escape(s: str) -> str:
    """Escape LaTeX-special characters in user-supplied text."""
    if not s:
        return ""
    # Process backslash first so we don't double-escape the replacement text.
    out = []
    for ch in s:
        if ch in _LATEX_ESCAPES:
            out.append(_LATEX_ESCAPES[ch])
        elif ch in _UNICODE_TEXT_MAP:
            out.append(_UNICODE_TEXT_MAP[ch])
        else:
            out.append(ch)
    return "".join(out)
