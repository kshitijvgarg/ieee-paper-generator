# IEEE Paper Generator

Python tool that renders a YAML paper spec into IEEE-formatted `.docx`
(using the official IEEE Word template) or `.tex` (using IEEEtran), which
then compiles to PDF. One input spec, two output formats, matching IEEE
Conference A4 formatting in both.

## Layout

```
ieee-paper-generator/
├── ieee_paper/             Python package: YAML → .docx / .tex
│   └── examples/           Sample paper specs (YAML + optional .md bodies)
├── templates/
│   ├── docx/               IEEE Word template (consumed by the .docx builder)
│   └── latex/              Official IEEE LaTeX bundle (reference material)
├── build/                  Generated outputs (gitignored)
├── requirements.txt        Python dependencies
└── .venv/                  Virtual environment (gitignored; see setup)
```

## Quick start

Install Python deps once (from this folder):

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Render a paper from a YAML spec:

```bash
# DOCX output (opens in Word / LibreOffice / Pages)
.venv/bin/python -m ieee_paper build ieee_paper/examples/sample_paper.yml \
    -o build/paper.docx

# DOCX → PDF (headless LibreOffice; outputs alongside the .docx)
soffice --headless --convert-to pdf --outdir build build/paper.docx

# LaTeX output (compile with pdflatex)
.venv/bin/python -m ieee_paper build ieee_paper/examples/sample_paper.yml \
    -o build/paper.tex
pdflatex -output-directory=build build/paper.tex
```

Format is inferred from the output extension; use `--format docx|tex` to
override.

## Using from another project

To render a YAML living outside this repo, point `PYTHONPATH` at this repo
and invoke its venv's Python. No install step; the repo acts as a tool:

```bash
GEN=/path/to/ieee-paper-generator
mkdir -p out

# YAML → DOCX → PDF
PYTHONPATH=$GEN $GEN/.venv/bin/python -m ieee_paper build paper.yml -o out/paper.docx \
  && soffice --headless --convert-to pdf --outdir out out/paper.docx

# YAML → TEX → PDF
PYTHONPATH=$GEN $GEN/.venv/bin/python -m ieee_paper build paper.yml -o out/paper.tex \
  && pdflatex -output-directory=out out/paper.tex
```

Paths in the YAML (e.g. figure `path`, `body` markdown files) resolve relative
to the spec's directory, so referenced assets stay colocated with your paper.

## LaTeX prerequisites

The LaTeX path needs a TeX distribution with `IEEEtran` and `multirow`. On
macOS with BasicTeX:

```bash
eval "$(/usr/libexec/path_helper)"          # refresh shell PATH once
sudo tlmgr update --self
sudo tlmgr install ieeetran multirow
```

After that, `pdflatex <file>.tex` works without env-var tricks.

## YAML spec

- [ieee_paper/examples/sample_paper.yml](ieee_paper/examples/sample_paper.yml)
  — minimal working example.
- [ieee_paper/examples/template_reproduction.yml](ieee_paper/examples/template_reproduction.yml)
  — full reproduction of the blank IEEE template, exercising every supported
  feature: sections at all levels, bullet lists, equations, tables with
  cell merges (`rowspan`/`colspan`), figure captions, sponsor footnote
  frame, and numbered references.
