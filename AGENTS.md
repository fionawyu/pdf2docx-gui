# Repository Guidelines

## Project Structure & Modules

This Python desktop app keeps its tracked code at the repository root. `pdf2docx_gui.py` starts the GUI and PDF CLI; `fluent_ui.py` holds the Fluent pages; `latex_to_docx.py` converts LaTeX with Pandoc. `snip_regions.py` crops PDF tables and equations, `polish.py` repairs converted documents, and `check.py` compares DOCX content with its PDF. Dependencies are in `requirements.txt`; `run_windows.bat` and `run_macos.sh` launch the app. Regression tests are in `tests/`.

## Build, Run & Validate

There is no build step. In a virtual environment, run `pip install -r requirements.txt`, then `python pdf2docx_gui.py`. Launchers create the environment and repair missing Fluent Widgets or Pandoc dependencies. PDF scripting uses `python pdf2docx_gui.py --cli input.pdf [output.docx] [--no-images] [--no-tidy]`. The LaTeX page writes a sibling `_from_tex.docx` and uses the source folder for figures and declared bibliographies. Check PDF output with `python check.py input.pdf output.docx`.

## Coding Style & Naming

Use Python with four spaces per indentation level. Follow the existing `snake_case` naming for modules, functions, and variables; use descriptive names for conversion and document-repair logic. Keep dependencies in `requirements.txt` and maintain compatibility with the Python versions stated in `README.md`. No formatter or linter is configured, so keep edits consistent with nearby code and review them before submitting.

## Testing Guidelines

Tests use the standard-library `unittest` runner; run `python -m unittest discover -s tests -v` (GUI tests use Qt's offscreen platform). Name new test modules `test_*.py`. For PDF changes, run the CLI and `check.py` on a representative document. For LaTeX changes, verify extracted text, editable tables/math, and embedded figures in the output DOCX. For GUI changes, include a brief manual check of the relevant interaction.

## Commits & Pull Requests

Recent commits use short Conventional Commit prefixes such as `feat:` and `docs:`. Use a focused message, for example `fix: preserve equation crops`. Pull requests should explain the user-visible change, note validation performed, and link a related issue when one exists. Include screenshots for visible GUI changes and a sample PDF/DOCX description for conversion changes; do not commit private source documents or generated outputs.

## Configuration & Local Files

Conversion runs locally and needs no credentials. Keep virtual environments and generated `.docx` outputs out of commits; `.gitignore` already excludes them.

## LaTeX Output Requirements

The TEX to DOCX output must look like the LaTeX PDF and be usable as a final deliverable without manual cleanup. Keep these behaviours when changing `latex_to_docx.py` (README "TEX to DOCX" lists them with details):

- Black text everywhere: no Word theme colours on title or headings, no blue or underlined links; Times New Roman, justified body, 1.5 spacing.
- Title block: one centered author line with superscript affiliation numbers, followed by the `authblk` affiliations.
- Caption labels in bold with LaTeX's separator (`Figure 1:` by default).
- Multi-panel (subfigure) figures as one pdflatex-rendered picture; tables and display equations as pdflatex-rendered pictures by default (GUI checkbox / CLI `--editable` for editable output).
- natbib papers: citations and references from the `.bbl`, reference list placed where `\bibliography` sits.
- No line may start with punctuation that follows a citation or cross-reference.
- Default output: `docx/<name>_from_tex.docx` next to the source.

