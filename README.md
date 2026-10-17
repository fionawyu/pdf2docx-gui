# PDF and LaTeX to DOCX Converter

[![Python](https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-blue)](#)
[![License](https://img.shields.io/badge/License-GPL--3.0-blue)](#license)

A cross-platform **desktop GUI** with separate PDF-to-Word and LaTeX-to-Word
pages. It converts PDF files into editable **Word (.docx)** documents and
converts LaTeX sources with their figures and bibliographies.

No file leaves your machine — everything is converted locally.

---

## Features

- Choose **PDF to Word** or **LaTeX to Word** from the Fluent side navigation.
- Drag and drop a source file onto its page, or use the Browse button.
- Choose the output `.docx` location.
- Keep PDF tables and equations as images for layout fidelity.
- Convert supported LaTeX tables and equations to editable Word content
  with Pandoc.
- Run conversions in the background while the Fluent UI stays responsive.
- Follow the system light or dark theme.
- Use the PDF `--cli` option for scripting without a window.
- Convert locally; no file leaves your machine.

## Tables and equations as images

A PDF holds positioned glyphs, not formulas. A converter has to guess where a
summation's limits or a fraction's numerator belong, and when it guesses wrong
the equation comes apart — taking any table built around one with it.

With **Tables and equations as images** ticked (the default), those regions are
cropped from the PDF at 300 dpi and placed in the document, so they look exactly
as they do in the PDF. Everything else is untouched: body text stays text, and
so do captions, which sit outside the crop and remain searchable.

Regions are found on the page itself — tables by the rules that frame them,
equations by being set apart from the margin and mostly set in math fonts — and
each crop replaces the converted content it actually came from, matched by text.

Turn it off when the tables matter more than the formulas do. **A journal
submission usually needs editable tables**, so check the instructions for
authors before sending a file with images in place of tables.

---

## Requirements

- **Python 3.9+** (3.12 recommended) and `pip`
- PySide6 Fluent Widgets (installed from `requirements.txt`)
- Pandoc, bundled by `pypandoc_binary` (about 41 MB for the Windows x64 wheel)

The PySide6 Fluent Widgets branch is used under its upstream GPLv3 terms;
commercial use of that toolkit requires its separate commercial license.
Pandoc is distributed under GPL-2.0-or-later; see its [license and source](https://github.com/jgm/pandoc/blob/main/README.md#license).

## Installation & Run

### Windows

Double-click **`run_windows.bat`**. It creates a virtual environment and
installs dependencies on first run, or when Fluent Widgets or Pandoc is missing
from an existing environment.

### macOS / Linux

```bash
chmod +x run_macos.sh
./run_macos.sh
```

### Manual (any OS)

```bash
python -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python pdf2docx_gui.py
```

> 💡 You can also pass a PDF as a command-line argument to preload it:
> `python pdf2docx_gui.py path/to/file.pdf`

---

## Usage

1. **Launch** the app and choose **PDF to Word** or **TEX to DOCX**
   from the side navigation.
2. On **PDF to Word**, drop a PDF onto the page (or click **Browse PDF**).
3. Confirm the output path (defaults to `<source-name>.docx`).
4. Decide whether to keep tables and equations as images (on by default).
5. Click **Convert**, then **Open DOCX** to open the saved file in your default app.

### TEX to DOCX

1. Select **TEX to DOCX** and choose or drop the main `.tex` source file.
2. Keep its figures and bibliography in the source folder at their referenced paths.
   For natbib papers, compile once with `latexmk` first so the `.bbl` is current.
3. Confirm the output (defaults to `docx/<source-name>_from_tex.docx` next to the
   source; the folder is created) and click **Convert**.
4. **Insert tables and equations as images** is on by default. Uncheck it for
   editable Word tables and equations.
5. Watch **Stages completed**, then click **Open DOCX**.

The result is meant to be a final deliverable that looks like the LaTeX PDF:

| Element | Word output |
|---|---|
| Title block | Centered title (not bold); authors on one line with superscript affiliation numbers; `authblk` affiliations below |
| Text | 12 pt Times New Roman, 1.5 spacing, justified; headings, title, citations and cross-references black (no Word theme blue, no link underline) |
| Captions | Bold label with LaTeX's separator, e.g. **Figure 1:** / **Table 1:** (`labelsep=period` keeps "Figure 1.") |
| Figures | Original image files; a figure with subfigure panels (a)(b)(c) becomes one picture rendered by `pdflatex`, exactly as in the PDF |
| Tables, display equations | Rendered by `pdflatex` and inserted as cropped pictures (default), or editable Word tables/equations when unchecked |
| Citations, references | natbib papers reuse LaTeX's `.bbl`: in-text "(Wang et al., 2023, 2024)", reference list in the PDF's style and order with a hanging indent, placed where `\bibliography` is (e.g. after Funding, before the appendices) |
| Line breaks | Punctuation after a citation or cross-reference stays on the same line |

Rendering tables, equations and multi-panel figures needs a LaTeX installation
(`pdflatex`). Without it, or if one item fails to compile, that item stays
editable and the page lists it under **Saved with warnings**.

If Pandoc can save a usable document while reporting missing figures, tables,
or references, the page shows **Saved with warnings** and lists the affected
content or resource paths. Review those items in the DOCX before using it.

### From a script

```bash
python pdf2docx_gui.py --cli input.pdf [output.docx] [--no-images] [--no-tidy]
```

`--no-images` converts tables and equations as text, as the tool did before.
`--no-tidy` skips the document cleanup pass.

```bash
python pdf2docx_gui.py --cli paper.tex [output.docx] [--editable]
```

Converts a LaTeX source with the same defaults as the **TEX to DOCX** page and
saves to `docx/paper_from_tex.docx`. `--editable` keeps tables and equations
editable instead of inserting them as pictures.

## Checking the result

A conversion can fail quietly: text is dropped, a caption disappears into a
cropped table, and nothing raises an error. `check.py` reads the PDF a page at
a time and reports what became of that page's words.

```bash
python check.py source.pdf converted.docx
```

It compares content, not layout, so it needs only the two files. Run it after
every conversion; it exits non-zero when anything is missing, so it can gate a
batch.

---

## Project layout

```text
pdf2docx-gui/
├── pdf2docx_gui.py      # GUI/CLI entry point and conversion pipeline
├── fluent_ui.py         # Fluent side navigation and conversion pages
├── latex_to_docx.py    # LaTeX-to-DOCX conversion using Pandoc
├── snip_regions.py      # finds tables and equations, crops them from the PDF
├── polish.py            # fonts, page flow, lost spaces, alignment
├── check.py             # compares the result against the PDF, page by page
├── requirements.txt     # dependencies
├── run_windows.bat      # Windows launcher (auto venv setup)
├── run_macos.sh         # macOS/Linux launcher (auto venv setup)
├── tests/               # focused regression tests
└── README.md            # this file
```

---

## How it works

The tool uses [`pdf2docx`](https://github.com/ArtifexSoftware/pdf2docx),
which is built on **PyMuPDF** (rendering & text extraction) and
**python-docx** (Word document generation). It parses each PDF page into
text blocks, vector shapes, and embedded images, then reassembles them
into a native `.docx`.

`snip_regions.py` then looks at the PDF again for the regions that survive
conversion badly, and replaces each with a crop of the page. `polish.py`
repairs what conversion page by page leaves behind — substituted fonts, the
per-page sections, the spaces lost from justified lines.

---

The LaTeX page calls Pandoc with the selected source folder as its resource
path and processes declared `.bib` files with citeproc. It does not compile or
modify the LaTeX source.

---

## Known limitations

- Complex vector graphics and exotic embedded fonts may render
  approximately (reflowed or substituted) rather than pixel-perfect.
- Very large, image-heavy PDFs can take a while; the progress bar shows
  activity but not a percentage.
- Scanned/image-only PDFs will not produce selectable text (OCR is not
  included).
- A cropped table or equation is a picture: it does not reflow, and it cannot
  be edited as a Word table or equation. Its extracted text is stored in the
  image description for accessibility and content checking.
- A table that runs across a PDF page boundary is cropped once per page, so a
  repeated header row appears twice.
- Where the text reflows, page breaks no longer fall where the PDF put them.
  A break is kept only where the PDF page stopped early or the next page opens
  a section.

### What goes wrong, and what the tool does about it

Every one of these was a real defect found by comparing output against the
source page by page. They are recorded because the next PDF will have them too.

| What goes wrong | Why | What the tool does |
| --- | --- | --- |
| Text looks a size too large | The .docx names TeX fonts nobody has installed, so Word substitutes something wider | Rewrites them to Times New Roman and Courier New |
| A blank page after almost every page | Each PDF page becomes its own section, ending with the page number as body text; reflowed text runs a little long and spills | Moves the page number into a real footer and drops the per-page sections |
| A gap in the middle of a sentence | The paragraph that carried a section break stays behind as an empty paragraph | Removes empty paragraphs left between text |
| Whole lines run together as one word | Justified lines are set with squeezed spacing, and below a threshold the extractor reads the gap as no gap | Restores the space between runs, judged against the document's own vocabulary and any `.tex`/`.bbl` beside the PDF |
| The right margin is ragged | pdf2docx reads alignment off where the ink falls, so body text arrives left-aligned | Justifies paragraphs long enough to wrap, and only those |
| A heading stretched across the page | It shares a paragraph with the text under it, and justification pulls its single line apart | Leaves any paragraph holding a line break unjustified |
| A figure split across two pages | Nothing holds its panels and caption together once the text flows | Keeps panels, captions and images with what follows |
| The paper starts on the title page | Removing the per-page sections removed the deliberate breaks with the rest | Keeps a break where the PDF page ended early, or the next page opens a section |
| A caption vanishes into a crop | It repeats the wording of the table it describes, so it scores as part of it | Holds captions out of the match by name |
| A paragraph beside a table disappears | It named the same methods as the table, and overlap alone was taken as belonging | Requires one text to account for nearly all of the other, and trusts a converted table row over a sentence |

---

## License

Copyright (c) 2026 lc700x. This project is licensed under GNU GPLv3-only; see
[LICENSE](./LICENSE) for the complete terms.
