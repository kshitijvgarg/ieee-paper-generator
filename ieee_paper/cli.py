from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .builder import DEFAULT_TEMPLATE, build_paper
from .latex_builder import build_paper_latex
from .spec import PaperSpec


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m ieee_paper",
        description="Render a YAML paper spec to an IEEE-formatted .docx or .tex file.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    build = sub.add_parser("build", help="Build a paper from a spec")
    build.add_argument("spec", type=Path, help="Path to a paper spec YAML file")
    build.add_argument("-o", "--output", type=Path, required=True, help="Output path (.docx or .tex)")
    build.add_argument(
        "--format",
        choices=["auto", "docx", "tex"],
        default="auto",
        help="Output format (default: inferred from output extension)",
    )
    build.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE,
        help=f"Path to IEEE .docx template (docx only; default: {DEFAULT_TEMPLATE.name})",
    )

    args = parser.parse_args(argv)

    if args.command == "build":
        spec = PaperSpec.from_yaml(args.spec)
        fmt = _resolve_format(args.format, args.output)
        if fmt == "docx":
            build_paper(spec, args.output, template_path=args.template)
        elif fmt == "tex":
            build_paper_latex(spec, args.output)
        else:
            parser.error(f"Cannot infer format from output path '{args.output}'. "
                         "Use --format docx|tex.")
        print(f"Wrote {args.output}")
        return 0

    parser.print_help()
    return 1


def _resolve_format(flag: str, output: Path) -> str:
    if flag != "auto":
        return flag
    ext = output.suffix.lower()
    if ext == ".docx":
        return "docx"
    if ext in (".tex", ".latex"):
        return "tex"
    return ""


if __name__ == "__main__":
    sys.exit(main())
