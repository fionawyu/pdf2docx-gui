import os
import sys

from pdf2docx import Converter

import polish as polish_module
import snip_regions

SNIP_DPI = 300


def _vocab_beside(src):
    """Source files next to the PDF that name the words it uses."""
    stem = os.path.splitext(src)[0]
    return [path for path in (stem + ".tex", stem + ".bbl") if os.path.exists(path)]


def _convert_document(src, dst, snip=True, tidy=True, on_status=None):
    """Convert one PDF and return optional image-crop and polish results."""
    converter = Converter(src)
    try:
        converter.convert(dst, start=0, end=None)
    finally:
        converter.close()

    result = {"images": None, "polish": None}
    if snip:
        if on_status:
            on_status("Cropping tables and equations…")
        result["images"] = snip_regions.snip(
            src, dst, dpi=SNIP_DPI, tables=True, equations=True)
    if tidy:
        if on_status:
            on_status("Tidying the document…")
        result["polish"] = polish_module.polish(dst, src, _vocab_beside(src))
    return result


def build_app():
    """Create the Qt application and Fluent converter window."""
    from PySide6.QtWidgets import QApplication

    from fluent_ui import ConverterWindow

    app = QApplication.instance() or QApplication(sys.argv[:1])
    app.setApplicationName("PDF and LaTeX to Word Converter")
    return app, ConverterWindow(_convert_document)


def convert_cli(src, dst, snip=True, tidy=True):
    """Convert without opening the window, for scripting and for testing."""
    result = _convert_document(src, dst, snip=snip, tidy=tidy)
    if snip:
        replaced, missed = result["images"]
        print(f"{replaced} tables and equations kept as images"
              + (f", {missed} left as converted text" if missed else ""))
    if tidy:
        counts = result["polish"]
        print(", ".join(f"{name}: {value}" for name, value in counts.items()))
    print(dst)


def convert_tex_cli(src, dst=None, as_images=True):
    """Convert a .tex file the way the LaTeX page does, printing any warnings."""
    from fluent_ui import _default_tex_output
    from latex_to_docx import convert_latex_to_docx

    dst = dst or _default_tex_output(src)
    warnings = []
    convert_latex_to_docx(src, dst, on_warning=warnings.append,
                          tables_as_images=as_images, equations_as_images=as_images)
    for message in warnings:
        print("warning:", message)
    print(dst)
    return 0


def main():
    args = [arg for arg in sys.argv[1:] if not arg.startswith("-")]
    flags = {arg for arg in sys.argv[1:] if arg.startswith("-")}

    if "--cli" in flags:
        if not args:
            print("usage: pdf2docx_gui.py --cli input.pdf [output.docx] "
                  "[--no-images] [--no-tidy]\n"
                  "       pdf2docx_gui.py --cli paper.tex [output.docx] [--editable]")
            return 2
        src = os.path.abspath(args[0])
        if src.lower().endswith(".tex"):
            return convert_tex_cli(src, os.path.abspath(args[1]) if len(args) > 1 else None,
                                   as_images="--editable" not in flags)
        dst = (os.path.abspath(args[1]) if len(args) > 1 else
               os.path.splitext(src)[0] + ".docx")
        convert_cli(src, dst,
                    snip="--no-images" not in flags,
                    tidy="--no-tidy" not in flags)
        return 0

    app, window = build_app()
    if args and args[0].lower().endswith(".pdf"):
        window.set_input(os.path.abspath(args[0]))
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
