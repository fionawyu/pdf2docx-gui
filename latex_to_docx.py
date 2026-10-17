"""Convert a LaTeX source document to an editable Word document with Pandoc."""

import json
import os
import re
import subprocess
import tempfile
from copy import deepcopy
from difflib import SequenceMatcher

import pypandoc
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt
from docx.text.paragraph import Paragraph


_BIBLIOGRAPHY_PATTERN = re.compile(
    r"\\(?:bibliography\b|addbibresource\b)\s*"
    r"(?:\[[^\]]*\]\s*)?\{([^}]+)\}"
)
_ENVIRONMENT_PATTERN = re.compile(r"\\(begin|end)\s*\{([^}]+)\}")
_CAPTION_PATTERN = re.compile(r"\\caption(?![A-Za-z])(\*)?")
_LABEL_PATTERN = re.compile(r"\\label\s*\{([^}]+)\}")
_SECTION_COMMAND_PATTERN = re.compile(
    r"\\(subparagraph|subsubsection|subsection|paragraph|chapter|section|part)"
    r"(\*)?(?![A-Za-z])")

_PAPER_SIZES = {
    "a4paper": (Mm(210), Mm(297)),
    "a5paper": (Mm(148), Mm(210)),
    "letterpaper": (Inches(8.5), Inches(11)),
    "legalpaper": (Inches(8.5), Inches(14)),
    "executivepaper": (Mm(184.15), Mm(266.7)),
}
_LENGTH_PATTERN = re.compile(r"^\s*(-?\d+(?:\.\d+)?)\s*(in|inch|cm|mm|pt|bp)?\s*$", re.I)


def _strip_tex_comments(text):
    """Remove TeX comments while preserving escaped percent signs and lines."""
    output = []
    index = 0
    while index < len(text):
        if text[index] == "%":
            slashes = 0
            previous = index - 1
            while previous >= 0 and text[previous] == "\\":
                slashes += 1
                previous -= 1
            if slashes % 2 == 0:
                newline = text.find("\n", index)
                output.append(" ")
                index = len(text) if newline < 0 else newline
                continue
        output.append(text[index])
        index += 1
    return "".join(output)


def _length_in_inches(value):
    match = _LENGTH_PATTERN.match(value)
    if not match:
        return None
    amount = float(match.group(1))
    unit = (match.group(2) or "pt").lower()
    factors = {
        "in": 1,
        "inch": 1,
        "cm": 1 / 2.54,
        "mm": 1 / 25.4,
        "pt": 1 / 72,
        "bp": 1 / 72.27,
    }
    return amount * factors[unit]


def _layout_from_source(source_text):
    """Read common paper-size and margin settings from a LaTeX preamble."""
    active_source = _strip_tex_comments(source_text)
    document_class = re.search(
        r"\\documentclass(?:\s*\[([^\]]*)\])?\s*\{[^}]+\}",
        active_source,
    )
    options = (document_class.group(1).split(",")
               if document_class and document_class.group(1) else [])
    options = [option.strip().lower() for option in options]
    page_size = next(
        (_PAPER_SIZES[option] for option in options if option in _PAPER_SIZES),
        _PAPER_SIZES["letterpaper"],
    )
    margins = {side: 1.0 for side in ("left", "right", "top", "bottom")}

    geometry_chunks = []
    geometry_package = re.search(
        r"\\usepackage(?:\s*\[([^\]]*)\])?\s*\{[^}]*\bgeometry\b[^}]*\}",
        active_source,
    )
    if geometry_package and geometry_package.group(1):
        geometry_chunks.append(geometry_package.group(1))
    geometry_chunks.extend(re.findall(r"\\geometry\s*\{([^}]*)\}", active_source))

    for chunk in geometry_chunks:
        for option in chunk.split(","):
            key, separator, value = option.partition("=")
            key = key.strip().lower()
            value = value.strip()
            if not separator:
                if key in _PAPER_SIZES:
                    page_size = _PAPER_SIZES[key]
                continue
            inches = _length_in_inches(value)
            if inches is None:
                continue
            if key == "margin":
                margins.update({side: inches for side in margins})
            elif key in margins:
                margins[key] = inches
            elif key == "hmargin":
                margins["left"] = margins["right"] = inches
            elif key == "vmargin":
                margins["top"] = margins["bottom"] = inches

    return page_size, margins


def _set_style_font(style, size, bold=None, italic=None):
    style.font.name = "Times New Roman"
    style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic
    style._element.get_or_add_rPr().get_or_add_rFonts().set(
        qn("w:eastAsia"), "Times New Roman")


def _style_by_name(document, name):
    return next(style for style in document.styles if style.name == name)


def _use_black_text(style):
    """Drop Word theme colours and fonts so the style prints like LaTeX output."""
    rpr = style._element.get_or_add_rPr()
    fonts = rpr.get_or_add_rFonts()
    for attribute in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        fonts.attrib.pop(qn(attribute), None)
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(attribute), "Times New Roman")
    color = rpr.find(qn("w:color"))
    if color is None:
        color = OxmlElement("w:color")
        rpr.append(color)
    for attribute in ("w:themeColor", "w:themeShade", "w:themeTint"):
        color.attrib.pop(qn(attribute), None)
    color.set(qn("w:val"), "000000")


def _use_black_theme_styles(document):
    """Make titles, headings, and links black instead of Word's accent blue."""
    for style in document.styles:
        rpr = style._element.find(qn("w:rPr"))
        color = rpr.find(qn("w:color")) if rpr is not None else None
        uses_accent = (color is not None
                       and color.get(qn("w:themeColor"), "").startswith("accent"))
        if uses_accent or style.name in ("Title", "Hyperlink"):
            _use_black_text(style)
    try:
        _style_by_name(document, "Hyperlink").font.underline = False
    except StopIteration:
        pass


def _ensure_table_caption_style(document):
    try:
        return _style_by_name(document, "Table Caption")
    except StopIteration:
        style = document.styles.add_style("Table Caption", WD_STYLE_TYPE.PARAGRAPH)
        try:
            style.base_style = _style_by_name(document, "Caption")
        except StopIteration:
            style.base_style = _style_by_name(document, "Normal")
        _set_style_font(style, 10, italic=False)
        style.paragraph_format.line_spacing = 1
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.keep_with_next = True
        return style


def _create_reference_docx(source_text, destination_path):
    """Create a Pandoc reference file using the source paper's basic layout."""
    pandoc = pypandoc.get_pandoc_path()
    result = subprocess.run(
        [pandoc, "--print-default-data-file", "reference.docx"],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    with open(destination_path, "wb") as reference_file:
        reference_file.write(result.stdout)

    document = Document(destination_path)
    page_size, margins = _layout_from_source(source_text)
    for section in document.sections:
        section.page_width, section.page_height = page_size
        section.left_margin = Inches(margins["left"])
        section.right_margin = Inches(margins["right"])
        section.top_margin = Inches(margins["top"])
        section.bottom_margin = Inches(margins["bottom"])

    normal = _style_by_name(document, "Normal")
    _set_style_font(normal, 12)
    normal.paragraph_format.line_spacing = None

    for name in ("Body Text", "First Paragraph", "Bibliography", "Abstract"):
        try:
            style = _style_by_name(document, name)
        except StopIteration:
            continue
        _set_style_font(style, 12)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.space_before = None
        style.paragraph_format.space_after = Pt(6)
        if name != "Bibliography":
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    try:
        _style_by_name(document, "Abstract").paragraph_format.keep_with_next = True
    except StopIteration:
        pass
    try:
        abstract_title = _style_by_name(document, "Abstract Title")
        _set_style_font(abstract_title, 12, bold=True)
        abstract_title.paragraph_format.space_before = Pt(15)
        abstract_title.paragraph_format.space_after = None
        abstract_title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        abstract_title.paragraph_format.keep_with_next = True
    except StopIteration:
        pass
    try:
        compact = _style_by_name(document, "Compact")
        _set_style_font(compact, 12)
        compact.paragraph_format.space_before = Pt(1.8)
        compact.paragraph_format.space_after = Pt(1.8)
    except StopIteration:
        pass

    title = _style_by_name(document, "Title")
    _set_style_font(title, 17, bold=False)
    title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.line_spacing = 1
    for name, size in (("Heading 1", 14), ("Heading 2", 12), ("Heading 3", 12)):
        heading = _style_by_name(document, name)
        _set_style_font(heading, size, bold=True, italic=False)
        heading.paragraph_format.space_before = Pt(12)
        heading.paragraph_format.space_after = Pt(6)
    for name in ("Table Caption", "Image Caption"):
        try:
            caption = _style_by_name(document, name)
        except StopIteration:
            caption = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        _set_style_font(caption, 10, italic=False)
        caption.paragraph_format.line_spacing = 1
        caption.paragraph_format.space_before = None
        caption.paragraph_format.space_after = Pt(8)
    _style_by_name(document, "Table Caption").paragraph_format.keep_with_next = True

    for name in ("Author", "Date"):
        try:
            _set_style_font(_style_by_name(document, name), 12)
        except KeyError:
            pass
    _use_black_theme_styles(document)
    document.save(destination_path)


_AUTHOR_PATTERN = re.compile(r"\\author\s*(?:\[([^\]]*)\])?\s*\{([^{}]*)\}")
_AFFILIATION_PATTERN = re.compile(r"\\affil\s*(?:\[([^\]]*)\])?\s*\{([^{}]*)\}")


def _title_block_records(source_text):
    """Return (authors, affiliations) from authblk-style \\author/\\affil commands."""
    text = _strip_tex_comments(source_text)
    preamble = text.split(r"\begin{document}", 1)[0]
    authors = [(mark or "", name.strip()) for mark, name in _AUTHOR_PATTERN.findall(preamble)]
    affiliations = [(mark or "", name.strip())
                    for mark, name in _AFFILIATION_PATTERN.findall(preamble)]
    return authors, affiliations


def _clear_runs(paragraph):
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)


def _restore_title_block(source_text, output_path):
    """Put authors on one centered line and add authblk affiliations below them.

    Pandoc writes one Author paragraph per \\author and drops \\affil entirely.
    """
    authors, affiliations = _title_block_records(source_text)
    if len(authors) < 2 and not affiliations:
        return
    document = Document(output_path)
    author_paragraphs = [paragraph for paragraph in document.paragraphs
                         if paragraph.style.name == "Author"]
    if not author_paragraphs:
        return
    first = author_paragraphs[0]
    for paragraph in author_paragraphs[1:]:
        paragraph._element.getparent().remove(paragraph._element)
    _clear_runs(first)
    first.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for index, (mark, name) in enumerate(authors):
        if index:
            first.add_run(", and " if index == len(authors) - 1 and len(authors) > 2
                          else (" and " if index == len(authors) - 1 else ", "))
        first.add_run(name)
        if mark and affiliations:
            first.add_run(mark).font.superscript = True

    anchor = first
    for mark, name in affiliations:
        element = deepcopy(first._element)
        anchor._element.addnext(element)
        paragraph = Paragraph(element, first._parent)
        _clear_runs(paragraph)
        if mark:
            paragraph.add_run(mark).font.superscript = True
        paragraph.add_run(name).font.size = Pt(11)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        anchor = paragraph
    anchor.paragraph_format.space_after = Pt(12)
    document.save(output_path)


def _is_comment_offset(text, offset):
    line_start = text.rfind("\n", 0, offset) + 1
    index = line_start
    while index < offset:
        if text[index] == "%":
            slash_count = 0
            previous = index - 1
            while previous >= line_start and text[previous] == "\\":
                slash_count += 1
                previous -= 1
            if slash_count % 2 == 0:
                return True
        index += 1
    return False


def _wide_subfigure_ranges(text):
    """Find subfigures whose image width should follow a wide line width."""
    environments = [
        match for match in _ENVIRONMENT_PATTERN.finditer(text)
        if not _is_comment_offset(text, match.start())
    ]
    ranges = []
    for index, match in enumerate(environments):
        if match.groups() != ("begin", "subfigure"):
            continue
        cursor = match.end()
        if cursor < len(text) and text[cursor] == "[":
            _, cursor = _read_tex_group(text, cursor, "[", "]")
        width, _ = _read_tex_group(text, cursor, "{", "}")
        ratio = re.fullmatch(
            r"\s*(?:(\d+(?:\.\d+)?)\s*)?\\(?:textwidth|linewidth|columnwidth)\s*",
            width or "")
        ratio_value = float(ratio.group(1) or 1) if ratio else 0
        if not ratio or ratio_value < 0.5:
            continue

        depth = 1
        for closing in environments[index + 1:]:
            if closing.group(2) != "subfigure":
                continue
            depth += 1 if closing.group(1) == "begin" else -1
            if not depth:
                ranges.append((match.start(), closing.end(), str(ratio_value)))
                break
    return ranges


def _wide_subfigure_linewidth_images(text):
    """Map source image positions to widths lost by Pandoc's subfigure reader."""
    subfigures = _wide_subfigure_ranges(text)
    widths = []
    image_index = 0
    for match in re.finditer(r"\\includegraphics\*?", text):
        if _is_comment_offset(text, match.start()):
            continue
        cursor = match.end()
        options = None
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        if cursor < len(text) and text[cursor] == "[":
            options, _ = _read_tex_group(text, cursor, "[", "]")
        containing = [item for item in subfigures
                      if item[0] < match.start() < item[1]]
        if (options and containing
                and re.search(
                    r"\bwidth\s*=\s*\\(?:textwidth|linewidth|columnwidth)\b",
                    options)):
            width = max(containing, key=lambda item: item[0])[2]
            widths.append((image_index, float(width)))
        image_index += 1
    return image_index, widths


def _normalize_includegraphics_options(text):
    """Flatten commented, multiline includegraphics options for Pandoc's reader."""
    pattern = re.compile(r"\\includegraphics\*?\s*\[")
    output = []
    cursor = 0
    for match in pattern.finditer(text):
        if match.start() < cursor or _is_comment_offset(text, match.start()):
            continue

        options = []
        square_depth = 1
        brace_depth = 0
        index = match.end()
        while index < len(text):
            character = text[index]
            if character == "%":
                slash_count = 0
                previous = index - 1
                while previous >= match.end() and text[previous] == "\\":
                    slash_count += 1
                    previous -= 1
                if slash_count % 2 == 0:
                    newline = text.find("\n", index)
                    if newline < 0:
                        index = len(text)
                        break
                    options.append(" ")
                    index = newline + 1
                    continue
            if character == "\\" and index + 1 < len(text):
                options.append(text[index:index + 2])
                index += 2
                continue
            if character == "{":
                brace_depth += 1
            elif character == "}" and brace_depth:
                brace_depth -= 1
            elif not brace_depth and character == "[":
                square_depth += 1
            elif not brace_depth and character == "]":
                square_depth -= 1
                if not square_depth:
                    break
            options.append(" " if character in "\r\n\t" else character)
            index += 1

        if square_depth:
            continue
        output.extend((text[cursor:match.start()], text[match.start():match.end()],
                       "".join(options).strip(), "]"))
        cursor = index + 1

    if not output:
        return text
    output.append(text[cursor:])
    return "".join(output)


def _read_tex_group(text, index, opening, closing):
    while index < len(text) and text[index].isspace():
        index += 1
    if index >= len(text) or text[index] != opening:
        return None, index

    depth = 1
    content = []
    index += 1
    while index < len(text) and depth:
        character = text[index]
        if character == "%":
            slashes = 0
            previous = index - 1
            while previous >= 0 and text[previous] == "\\":
                slashes += 1
                previous -= 1
            if slashes % 2 == 0:
                newline = text.find("\n", index)
                content.append(" ")
                index = len(text) if newline < 0 else newline
                continue
        if character == "\\" and index + 1 < len(text):
            content.append(text[index:index + 2])
            index += 2
            continue
        if character == opening:
            depth += 1
        elif character == closing:
            depth -= 1
            if not depth:
                break
        content.append(character)
        index += 1
    if depth:
        return None, index
    return "".join(content), index + 1


def _table_float_ranges(text):
    ranges = []
    stack = []
    for match in _ENVIRONMENT_PATTERN.finditer(text):
        if _is_comment_offset(text, match.start()):
            continue
        operation, environment = match.groups()
        if operation == "begin":
            if environment in ("table", "table*") and not any(
                    entry[0] in ("table", "table*") for entry in stack):
                stack.append((environment, match.start()))
            else:
                stack.append((environment, None))
        elif stack:
            matching = next((index for index in range(len(stack) - 1, -1, -1)
                             if stack[index][0] == environment), None)
            if matching is None:
                continue
            open_environment, start = stack[matching]
            del stack[matching:]
            if open_environment in ("table", "table*") and start is not None:
                ranges.append((start, match.end()))
    return ranges


def _named_environment_ranges(text, names):
    ranges = []
    stack = []
    for match in _ENVIRONMENT_PATTERN.finditer(text):
        if _is_comment_offset(text, match.start()):
            continue
        operation, environment = match.groups()
        if operation == "begin":
            start = match.start() if environment in names and not any(
                item[0] in names and item[1] is not None for item in stack) else None
            stack.append((environment, start))
        elif stack:
            matching = next((index for index in range(len(stack) - 1, -1, -1)
                             if stack[index][0] == environment), None)
            if matching is None:
                continue
            open_environment, start = stack[matching]
            del stack[matching:]
            if open_environment in names and start is not None:
                ranges.append((start, match.end()))
    return ranges


def _figure_caption_records(text):
    """Return numbered figure captions whose labels Word/Pandoc omits."""
    records = []
    figure_number = 0
    for start, end in _named_environment_ranges(text, {"figure", "figure*"}):
        block = text[start:end]
        subfigures = _named_environment_ranges(block, {"subfigure"})
        captions = [match for match in _CAPTION_PATTERN.finditer(block)
                    if not _is_comment_offset(block, match.start())]
        top_level = [match for match in captions
                     if not match.group(1)
                     and not any(left < match.start() < right
                                 for left, right in subfigures)]
        for match in top_level:
            figure_number += 1
            cursor = match.end()
            while cursor < len(block) and block[cursor].isspace():
                cursor += 1
            if cursor < len(block) and block[cursor] == "[":
                _, cursor = _read_tex_group(block, cursor, "[", "]")
            caption, _ = _read_tex_group(block, cursor, "{", "}")
            if caption is None:
                continue
            records.append({
                "caption": caption,
                "prefix": f"Figure {figure_number}.",
                "number": figure_number,
                "float": block,
            })
    return records


def _read_section_command(text, match):
    cursor = match.end()
    while cursor < len(text) and text[cursor].isspace():
        cursor += 1
    if cursor < len(text) and text[cursor] == "[":
        _, cursor = _read_tex_group(text, cursor, "[", "]")
    while cursor < len(text) and text[cursor].isspace():
        cursor += 1
    title, cursor = _read_tex_group(text, cursor, "{", "}")
    if title is None:
        return None

    labels = []
    while cursor < len(text):
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        label = _LABEL_PATTERN.match(text, cursor)
        if not label:
            break
        labels.append(label.group(1).strip())
        cursor = label.end()
    return title, cursor, labels


def _section_label_aliases(source_text):
    text = _strip_tex_comments(source_text)
    aliases = {}
    for match in _SECTION_COMMAND_PATTERN.finditer(text):
        record = _read_section_command(text, match)
        if record is None:
            continue
        _, _, labels = record
        if labels:
            aliases.update({label: labels[0] for label in labels[1:]})
    return aliases


def _appendix_heading_records(source_text):
    """Build the lettered heading sequence after LaTeX's appendix switch."""
    text = _strip_tex_comments(source_text)
    appendix = re.search(r"\\appendix\b", text)
    if not appendix:
        return [], {}

    level_by_command = {
        "section": 1, "subsection": 2, "subsubsection": 3,
    }
    records = []
    labels = {}
    appendix_number = 0
    appendix_letter = None
    counters = {2: 0, 3: 0}
    for match in _SECTION_COMMAND_PATTERN.finditer(text, appendix.end()):
        command, starred = match.groups()
        level = level_by_command.get(command)
        if level is None:
            continue
        parsed = _read_section_command(text, match)
        if parsed is None:
            continue
        title, _, heading_labels = parsed

        prefix = None
        if not starred:
            if level == 1:
                appendix_number += 1
                number = appendix_number
                letter = ""
                while number:
                    number, remainder = divmod(number - 1, 26)
                    letter = chr(ord("A") + remainder) + letter
                appendix_letter = letter
                counters = {2: 0, 3: 0}
                prefix = f"Appendix {appendix_letter}"
            elif appendix_letter is not None:
                counters[level] += 1
                for lower_level in range(level + 1, 4):
                    counters[lower_level] = 0
                suffix = ".".join(
                    str(counters[item]) for item in range(2, level + 1))
                prefix = f"{appendix_letter}.{suffix}"
        if prefix and heading_labels:
            number_text = prefix.removeprefix("Appendix ")
            labels.update({label: number_text for label in heading_labels})
        records.append({
            "level": level,
            "prefix": prefix,
            "title": title,
        })
    return records, labels


def _table_reference_numbers(source_text):
    numbers = {}
    for record in _table_float_records(source_text):
        if record["starred"]:
            continue
        active_float = _strip_tex_comments(record["float"])
        for label in _LABEL_PATTERN.findall(active_float):
            numbers[label.strip()] = str(record["number"])
    return numbers


def _figure_reference_numbers(source_text):
    numbers = {}
    for record in _figure_caption_records(_strip_tex_comments(source_text)):
        block = record["float"]
        subfigures = _named_environment_ranges(block, {"subfigure"})
        for match in _LABEL_PATTERN.finditer(block):
            if _is_comment_offset(block, match.start()):
                continue
            subfigure = next((index for index, (start, end) in enumerate(
                subfigures, start=1) if start < match.start() < end), None)
            number = str(record["number"])
            if subfigure is not None:
                number += chr(ord("a") + subfigure - 1)
            numbers[match.group(1).strip()] = number
    return numbers


def _table_float_records(text):
    """Return numbered table floats and their caption source text."""
    records = []
    table_number = 0
    for start, end in _table_float_ranges(text):
        block = text[start:end]
        for match in _CAPTION_PATTERN.finditer(block):
            if _is_comment_offset(block, match.start()):
                continue
            cursor = match.end()
            while cursor < len(block) and block[cursor].isspace():
                cursor += 1
            if cursor < len(block) and block[cursor] == "[":
                _, cursor = _read_tex_group(block, cursor, "[", "]")
            caption, caption_end = _read_tex_group(block, cursor, "{", "}")
            if caption is None:
                continue
            starred = bool(match.group(1))
            if not starred:
                table_number += 1
            records.append({
                "number": table_number,
                "caption": caption.strip(),
                "starred": starred,
                "float": block,
            })
    return records


def _normalize_table_captions(text):
    """Keep table captions as editable paragraphs even for unsupported floats."""
    replacements = []
    table_number = 0
    for start, end in _table_float_ranges(text):
        block = text[start:end]
        captions = []
        removals = []
        for match in _CAPTION_PATTERN.finditer(block):
            if _is_comment_offset(block, match.start()):
                continue
            cursor = match.end()
            while cursor < len(block) and block[cursor].isspace():
                cursor += 1
            if cursor < len(block) and block[cursor] == "[":
                _, cursor = _read_tex_group(block, cursor, "[", "]")
            caption, caption_end = _read_tex_group(block, cursor, "{", "}")
            if caption is None:
                continue
            starred = bool(match.group(1))
            if not starred:
                table_number += 1
                prefix = f"\\textbf{{Table {table_number}.}} "
            else:
                prefix = ""
            captions.append(prefix + caption.strip())
            removals.append((start + match.start(), start + caption_end))

        if captions:
            replacements.append((start, start, "\n\n" + "\n\n".join(captions) + "\n\n"))
            replacements.extend((begin, finish, "") for begin, finish in removals)

    for begin, finish, replacement in sorted(replacements, reverse=True):
        text = text[:begin] + replacement + text[finish:]
    return text


def _simplify_column_spec(spec):
    """Reduce tabularx column definitions to Word-friendly alignment columns."""
    columns = []
    index = 0
    alignment = None
    while index < len(spec):
        character = spec[index]
        if character.isspace() or character == "|":
            index += 1
            continue
        if spec.startswith(">{", index):
            directive, index = _read_tex_group(spec, index + 1, "{", "}")
            if directive and "\\centering" in directive:
                alignment = "c"
            elif directive and "\\raggedleft" in directive:
                alignment = "r"
            elif directive and "\\raggedright" in directive:
                alignment = "l"
            continue
        if character == "<":
            if index + 1 < len(spec) and spec[index + 1] == "{":
                _, index = _read_tex_group(spec, index + 1, "{", "}")
            else:
                index += 1
            continue
        if character in "pmb":
            _, index = _read_tex_group(spec, index + 1, "{", "}")
            columns.append(alignment or "l")
            alignment = None
            continue
        if character in "lcrXY":
            columns.append(alignment or {"X": "l", "Y": "l"}.get(character, character))
            alignment = None
            index += 1
            continue
        if character == "*":
            count, index = _read_tex_group(spec, index + 1, "{", "}")
            repeated_spec, index = _read_tex_group(spec, index, "{", "}")
            if count and repeated_spec:
                try:
                    columns.extend(_simplify_column_spec(repeated_spec) * int(count))
                except ValueError:
                    pass
            continue
        if character in "@!":
            _, index = _read_tex_group(spec, index + 1, "{", "}")
            continue
        if character == "\\":
            command = re.match(r"\\[A-Za-z]+", spec[index:])
            index += len(command.group(0)) if command else 2
            continue
        index += 1
    return "".join(columns) or "l"


def _normalize_table_markup(text):
    """Normalize table syntax once for Pandoc and standalone recovery passes."""
    for start, end in reversed(_named_environment_ranges(
            text, {"tabular", "tabularx"})):
        block = _strip_tex_comments(text[start:end])
        replacements = []
        for match in _ENVIRONMENT_PATTERN.finditer(block):
            if match.group(2) not in ("tabular", "tabularx"):
                continue
            operation, environment = match.groups()
            if operation == "end":
                if environment == "tabularx":
                    replacements.append((match.start(), match.end(), r"\end{tabular}"))
                continue

            cursor = match.end()
            if environment == "tabularx":
                _, cursor = _read_tex_group(block, cursor, "{", "}")
            spec, cursor = _read_tex_group(block, cursor, "{", "}")
            if spec is not None:
                replacements.append((
                    match.start(), cursor,
                    r"\begin{tabular}{" + _simplify_column_spec(spec) + "}"))

        for begin, finish, replacement in reversed(replacements):
            block = block[:begin] + replacement + block[finish:]

        command_replacements = []
        for match in re.finditer(r"\\makecell\b", block):
            if _is_comment_offset(block, match.start()):
                continue
            cursor = match.end()
            while cursor < len(block) and block[cursor].isspace():
                cursor += 1
            alignment = "c"
            if cursor < len(block) and block[cursor] == "[":
                options, cursor = _read_tex_group(block, cursor, "[", "]")
                align_match = re.search(r"[lcr]", options or "")
                if align_match:
                    alignment = align_match.group(0)
            body, finish = _read_tex_group(block, cursor, "{", "}")
            if body is not None:
                command_replacements.append((
                    match.start(), finish,
                    r"\begin{tabular}{" + alignment + "}"
                    + body + r"\end{tabular}"))

        for match in re.finditer(r"\\multicolumn\b", block):
            if _is_comment_offset(block, match.start()):
                continue
            cursor = match.end()
            groups = []
            for _ in range(3):
                group_start = cursor
                content, cursor = _read_tex_group(block, cursor, "{", "}")
                if content is None:
                    groups = []
                    break
                while group_start < len(block) and block[group_start].isspace():
                    group_start += 1
                groups.append((group_start, cursor, content))
            if groups:
                spec_start, spec_end, spec = groups[1]
                command_replacements.append((
                    spec_start, spec_end,
                    "{" + _simplify_column_spec(spec) + "}"))

        for begin, finish, replacement in sorted(command_replacements, reverse=True):
            block = block[:begin] + replacement + block[finish:]

        block = re.sub(
            r"\\(?:cmidrule|cline)\*?(?:\s*\[[^]]*\])?"
            r"(?:\([^)]*\))?\s*\{[^}]*\}", " ", block)
        block = re.sub(
            r"\\addlinespace\*?(?:\s*\[[^]]*\])?", " ", block)
        block = _normalize_tabular_rows(block)
        text = text[:start] + block + text[end:]
    return text


def _normalize_minipage_tables(text):
    """Flatten tabularx tables wrapped with a note minipage for Pandoc."""
    for start, end in reversed(_table_float_ranges(text)):
        block = text[start:end]
        active_environments = [
            match for match in _ENVIRONMENT_PATTERN.finditer(block)
            if not _is_comment_offset(block, match.start())
        ]
        if not any(match.groups() == ("begin", "minipage")
                   for match in active_environments):
            continue
        match = next((item for item in active_environments
                      if item.groups() == ("begin", "tabularx")), None)
        if match is None:
            continue
        index = match.end()
        _, index = _read_tex_group(block, index, "{", "}")
        spec, index = _read_tex_group(block, index, "{", "}")
        end_match = next((item for item in _ENVIRONMENT_PATTERN.finditer(block, index)
                          if item.groups() == ("end", "tabularx")
                          and not _is_comment_offset(block, item.start())), None)
        if spec is None or end_match is None:
            continue
        end_start = end_match.start()
        end_finish = end_match.end()
        begin = "\\begin{tabular}{" + _simplify_column_spec(spec) + "}"
        block = (block[:match.start()] + begin + block[index:end_start]
                 + "\\end{tabular}" + block[end_finish:])
        removals = []
        for environment_match in _ENVIRONMENT_PATTERN.finditer(block):
            if (_is_comment_offset(block, environment_match.start())
                    or environment_match.group(2) != "minipage"):
                continue
            if environment_match.group(1) == "begin":
                _, finish = _read_tex_group(
                    block, environment_match.end(), "{", "}")
            else:
                finish = environment_match.end()
            removals.append((environment_match.start(), finish))
        for begin, finish in reversed(removals):
            block = block[:begin] + block[finish:]
        text = text[:start] + block + text[end:]
    return text


def _normalize_tabular_rows(text):
    """Join table cells split across source lines before Pandoc reads them."""
    for start, end in reversed(_named_environment_ranges(
            text, {"tabular", "tabularx"})):
        block = text[start:end]
        lines = block.splitlines()
        normalized = []
        for line in lines:
            if line.lstrip().startswith("&"):
                previous = next((index for index in range(len(normalized) - 1, -1, -1)
                                 if normalized[index].strip()
                                 and not normalized[index].lstrip().startswith("%")), None)
                if (previous is not None
                        and not re.search(r"\\\\(?:\[[^]]*\])?\s*$",
                                          normalized[previous])):
                    normalized[previous] = normalized[previous].rstrip() + " " + line.strip()
                    continue
            normalized.append(line)
        text = text[:start] + "\n".join(normalized) + text[end:]
    return text


def _normalize_subfigure_tabularx(text):
    """Compatibility wrapper for the shared table normalization path."""
    return _normalize_table_markup(text)


def _standalone_table_sources(record, source_text):
    """Build standalone sources for each native table in a LaTeX float."""
    block = record["float"]
    document_start = re.search(r"\\begin\s*\{document\}", source_text)
    if not document_start:
        return []
    preamble = source_text[:document_start.start()]
    caption = record["caption"].strip()
    if not record["starred"]:
        caption = f"\\textbf{{Table {record['number']}.}} " + caption
    environments = [
        match for match in _ENVIRONMENT_PATTERN.finditer(block)
        if not _is_comment_offset(block, match.start())
    ]
    sources = []
    consumed_until = -1
    for index, match in enumerate(environments):
        environment = match.group(2)
        if (match.start() < consumed_until or match.group(1) != "begin"
                or environment not in ("tabularx", "tabular")):
            continue
        cursor = match.end()
        if environment == "tabularx":
            _, cursor = _read_tex_group(block, cursor, "{", "}")
        column_spec, cursor = _read_tex_group(block, cursor, "{", "}")
        if column_spec is None:
            continue

        depth = 1
        end_match = None
        for item in environments[index + 1:]:
            if item.start() < cursor or item.group(2) != environment:
                continue
            depth += 1 if item.group(1) == "begin" else -1
            if not depth:
                end_match = item
                break
        if end_match is None:
            continue

        table = _normalize_table_markup(block[match.start():end_match.end()])
        sources.append(
            preamble + "\n\\begin{document}\n" + caption + "\n\n"
            + table + "\n\\end{document}\n")
        consumed_until = end_match.end()
    return sources


def _table_column_specs(record):
    """Read width-bearing tabular column specifications from a float."""
    block = record["float"]
    environments = [match for match in _ENVIRONMENT_PATTERN.finditer(block)
                    if not _is_comment_offset(block, match.start())]
    specs = []
    for match in environments:
        if match.group(1) != "begin" or match.group(2) not in ("tabularx", "tabular"):
            continue
        cursor = match.end()
        table_width = None
        if match.group(2) == "tabularx":
            table_width, cursor = _read_tex_group(block, cursor, "{", "}")
        column_spec, _ = _read_tex_group(block, cursor, "{", "}")
        if column_spec is not None:
            specs.append((table_width, column_spec))
    return specs


def _column_width_expressions(spec):
    columns = []
    index = 0
    while index < len(spec):
        character = spec[index]
        if character.isspace() or character == "|":
            index += 1
            continue
        if spec.startswith(">{", index):
            _, index = _read_tex_group(spec, index + 1, "{", "}")
            continue
        if character == "<":
            if index + 1 < len(spec) and spec[index + 1] == "{":
                _, index = _read_tex_group(spec, index + 1, "{", "}")
            else:
                index += 1
            continue
        if character in "pmb":
            width, index = _read_tex_group(spec, index + 1, "{", "}")
            columns.append(width)
            continue
        if character in "lcrXY":
            columns.append(None)
            index += 1
            continue
        if character == "*":
            count, index = _read_tex_group(spec, index + 1, "{", "}")
            repeated, index = _read_tex_group(spec, index, "{", "}")
            if count and repeated:
                try:
                    columns.extend(_column_width_expressions(repeated) * int(count))
                except ValueError:
                    pass
            continue
        if character in "@!":
            _, index = _read_tex_group(spec, index + 1, "{", "}")
            continue
        if character == "\\":
            command = re.match(r"\\[A-Za-z]+", spec[index:])
            index += len(command.group(0)) if command else 2
            continue
        index += 1
    return columns


def _tex_width_inches(expression, reference_width):
    if expression is None:
        return None
    expression = expression.strip()
    relative = re.fullmatch(
        r"(-?(?:\d+(?:\.\d*)?|\.\d+))\s*\\(?:linewidth|textwidth|columnwidth)",
        expression,
    )
    if relative:
        return float(relative.group(1)) * reference_width
    if expression in (r"\linewidth", r"\textwidth", r"\columnwidth"):
        return reference_width
    return _length_in_inches(expression)


def _column_widths_inches(table_width_expression, column_spec, text_width):
    table_width = _tex_width_inches(table_width_expression, text_width)
    expressions = _column_width_expressions(column_spec)
    widths = [_tex_width_inches(value, text_width) for value in expressions]
    if not widths or not any(width is not None for width in widths):
        return None
    if table_width is None:
        return None

    fixed_total = sum(width for width in widths if width is not None)
    flexible_count = sum(width is None for width in widths)
    if fixed_total > table_width and fixed_total:
        scale = table_width / fixed_total
        widths = [width * scale if width is not None else None for width in widths]
        fixed_total = table_width
    if flexible_count:
        flexible_width = max(0.0, table_width - fixed_total) / flexible_count
        widths = [flexible_width if width is None else width for width in widths]
    elif fixed_total < table_width:
        scale = table_width / fixed_total
        widths = [width * scale for width in widths]
    return widths


def _set_table_column_widths(table_element, widths):
    grid = table_element.find(qn("w:tblGrid"))
    if grid is None:
        return False
    columns = list(grid.findall(qn("w:gridCol")))
    if len(columns) != len(widths):
        return False
    twips = [max(1, round(width * 1440)) for width in widths]
    table_properties = table_element.find(qn("w:tblPr"))
    if table_properties is None:
        table_properties = OxmlElement("w:tblPr")
        table_element.insert(0, table_properties)
    table_width = table_properties.find(qn("w:tblW"))
    if table_width is None:
        table_width = OxmlElement("w:tblW")
        table_properties.insert(0, table_width)
    table_width.set(qn("w:w"), str(sum(twips)))
    table_width.set(qn("w:type"), "dxa")
    layout = table_properties.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        table_properties.append(layout)
    layout.set(qn("w:type"), "fixed")

    for column, width in zip(columns, twips):
        column.set(qn("w:w"), str(width))
    for row in table_element.findall(qn("w:tr")):
        cell_widths = []
        column_index = 0
        for cell in row.findall(qn("w:tc")):
            cell_properties = cell.find(qn("w:tcPr"))
            span = (cell_properties.find(qn("w:gridSpan"))
                    if cell_properties is not None else None)
            span = int(span.get(qn("w:val"))) if span is not None else 1
            if column_index + span > len(twips):
                break
            cell_widths.append((cell, sum(twips[column_index:column_index + span])))
            column_index += span
        if column_index != len(twips):
            continue
        for cell, width in cell_widths:
            cell_properties = cell.find(qn("w:tcPr"))
            if cell_properties is None:
                cell_properties = OxmlElement("w:tcPr")
                cell.insert(0, cell_properties)
            cell_width = cell_properties.find(qn("w:tcW"))
            if cell_width is None:
                cell_width = OxmlElement("w:tcW")
                cell_properties.insert(0, cell_width)
            cell_width.set(qn("w:w"), str(width))
            cell_width.set(qn("w:type"), "dxa")
    return True


def _set_cell_padding(cell, padding_twips):
    cell_properties = cell.find(qn("w:tcPr"))
    if cell_properties is None:
        cell_properties = OxmlElement("w:tcPr")
        cell.insert(0, cell_properties)
    margins = cell_properties.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        cell_properties.append(margins)
    for side in ("top", "left", "bottom", "right"):
        margin = margins.find(qn(f"w:{side}"))
        if margin is None:
            margin = OxmlElement(f"w:{side}")
            margins.append(margin)
        margin.set(qn("w:w"), str(padding_twips))
        margin.set(qn("w:type"), "dxa")


def _cell_horizontal_padding(cell):
    cell_properties = cell.find(qn("w:tcPr"))
    margins = (cell_properties.find(qn("w:tcMar"))
               if cell_properties is not None else None)
    if margins is None:
        return 0
    return sum(
        int(margin.get(qn("w:w"), "0"))
        for side in ("left", "right")
        if (margin := margins.find(qn(f"w:{side}"))) is not None
    )


def _constrain_nested_table_widths(table_element):
    """Fit each nested table to its merged parent cell, preserving its ratios."""
    grid = table_element.find(qn("w:tblGrid"))
    if grid is None:
        return False
    widths = [int(column.get(qn("w:w"), "0"))
              for column in grid.findall(qn("w:gridCol"))]
    if not widths or not all(widths):
        return False

    changed = False
    for row in table_element.findall(qn("w:tr")):
        column_index = 0
        for cell in row.findall(qn("w:tc")):
            properties = cell.find(qn("w:tcPr"))
            span_node = (properties.find(qn("w:gridSpan"))
                         if properties is not None else None)
            span = int(span_node.get(qn("w:val"))) if span_node is not None else 1
            if column_index + span > len(widths):
                break
            available = max(
                1, sum(widths[column_index:column_index + span])
                - _cell_horizontal_padding(cell))
            column_index += span
            for nested in cell.findall(qn("w:tbl")):
                nested_grid = nested.find(qn("w:tblGrid"))
                nested_widths = (
                    [int(column.get(qn("w:w"), "0"))
                     for column in nested_grid.findall(qn("w:gridCol"))]
                    if nested_grid is not None else [])
                total = sum(nested_widths)
                if (nested_widths and all(nested_widths) and total > available):
                    scaled = [max(1, round(width * available / total))
                              for width in nested_widths]
                    scaled[-1] = max(1, scaled[-1] + available - sum(scaled))
                    changed = _set_table_column_widths(
                        nested, [width / 1440 for width in scaled]) or changed
                changed = _constrain_nested_table_widths(nested) or changed
    return changed


def _format_latex_tables(output_path):
    """Apply compact manuscript table typography without flattening emphasis."""
    document = Document(output_path)
    changed = False
    for table in document.element.iter(qn("w:tbl")):
        properties = table.find(qn("w:tblPr"))
        style = (properties.find(qn("w:tblStyle"))
                 if properties is not None else None)
        padding = 0 if style is not None and style.get(qn("w:val")) == "FigureTable" else 40
        for cell in table.findall(f".//{qn('w:tc')}"):
            _set_cell_padding(cell, padding)

        for paragraph in table.findall(f".//{qn('w:p')}"):
            paragraph_properties = paragraph.find(qn("w:pPr"))
            if paragraph_properties is None:
                paragraph_properties = OxmlElement("w:pPr")
                paragraph.insert(0, paragraph_properties)
            spacing = paragraph_properties.find(qn("w:spacing"))
            if spacing is None:
                spacing = OxmlElement("w:spacing")
                paragraph_properties.append(spacing)
            spacing.set(qn("w:before"), "0")
            spacing.set(qn("w:after"), "0")
            spacing.set(qn("w:line"), "240")
            spacing.set(qn("w:lineRule"), "auto")

        for run in table.iter(qn("w:r")):
            run_properties = run.find(qn("w:rPr"))
            if run_properties is None:
                run_properties = OxmlElement("w:rPr")
                run.insert(0, run_properties)
            for size_name in ("w:sz", "w:szCs"):
                size = run_properties.find(qn(size_name))
                if size is None:
                    size = OxmlElement(size_name)
                    run_properties.append(size)
                size.set(qn("w:val"), "20")

        rows = table.findall(f"./{qn('w:tr')}")
        header_rows = []
        for row in rows:
            row_properties = row.find(qn("w:trPr"))
            if row_properties is None:
                continue
            header = row_properties.find(qn("w:tblHeader"))
            if header is not None:
                header.set(qn("w:val"), "true")
                header_rows.append(row)
        if style is None or style.get(qn("w:val")) != "FigureTable":
            if rows and not header_rows:
                row_properties = rows[0].find(qn("w:trPr"))
                if row_properties is None:
                    row_properties = OxmlElement("w:trPr")
                    rows[0].insert(0, row_properties)
                row_properties.append(OxmlElement("w:tblHeader"))
        changed = True

    if changed:
        document.save(output_path)


def _keep_table_captions_with_tables(output_path):
    """Keep captions and their notes attached to the following native table."""
    document = Document(output_path)
    children = list(document.element.body)
    changed = False
    caption_pattern = re.compile(r"^Table\s+\d+\.")
    for index, element in enumerate(children):
        if element.tag != qn("w:p") or not caption_pattern.match(
                Paragraph(element, document).text.strip()):
            continue
        caption_properties = element.find(qn("w:pPr"))
        if caption_properties is None:
            caption_properties = OxmlElement("w:pPr")
            element.insert(0, caption_properties)
        if caption_properties.find(qn("w:keepNext")) is None:
            caption_properties.append(OxmlElement("w:keepNext"))
            changed = True
        for following in children[index + 1:]:
            if following.tag == qn("w:tbl"):
                break
            if following.tag == qn("w:p"):
                if caption_pattern.match(Paragraph(following, document).text.strip()):
                    break
                properties = following.find(qn("w:pPr"))
                if properties is None:
                    properties = OxmlElement("w:pPr")
                    following.insert(0, properties)
                keep_next = properties.find(qn("w:keepNext"))
                if keep_next is None:
                    properties.append(OxmlElement("w:keepNext"))
                    changed = True

    if changed:
        document.save(output_path)


def _constrain_document_nested_tables(output_path):
    document = Document(output_path)
    changed = False
    for element in document.element.body.findall(qn("w:tbl")):
        changed = _constrain_nested_table_widths(element) or changed
    if changed:
        document.save(output_path)


def _preserve_latex_table_widths(source_text, output_path):
    """Apply explicit LaTeX table widths to the corresponding editable Word tables."""
    document = Document(output_path)
    section = document.sections[0]
    text_width = (section.page_width - section.left_margin - section.right_margin) / 914400
    body = document.element.body
    children = list(body)
    changed = False

    for record in _table_float_records(source_text):
        if record["starred"]:
            continue
        source_widths = [
            widths for table_width, column_spec in _table_column_specs(record)
            if (widths := _column_widths_inches(
                table_width, column_spec, text_width)) is not None
        ]
        if not source_widths:
            continue
        marker = f"Table {record['number']}."
        caption_index = next((index for index, element in enumerate(children)
                              if element.tag == qn("w:p")
                              and Paragraph(element, document).text.strip().startswith(marker)),
                             None)
        if caption_index is None:
            continue

        tables = []
        for element in children[caption_index + 1:]:
            if element.tag == qn("w:p") and re.match(
                    r"^Table\s+\d+\.", Paragraph(element, document).text.strip()):
                break
            if (element.tag == qn("w:tbl")
                    and not element.findall(".//" + qn("w:drawing"))):
                tables.append(element)
        for table_element, widths in zip(tables, source_widths):
            changed = _set_table_column_widths(table_element, widths) or changed

    if changed:
        document.save(output_path)


def _subfigure_blocks(figure_record):
    blocks = []
    figure = figure_record["float"]
    for start, end in _named_environment_ranges(figure, {"subfigure"}):
        block = figure[start:end]
        begin = next((match for match in _ENVIRONMENT_PATTERN.finditer(block)
                      if match.groups() == ("begin", "subfigure")
                      and not _is_comment_offset(block, match.start())), None)
        if begin is None:
            continue
        cursor = begin.end()
        if cursor < len(block) and block[cursor] == "[":
            _, cursor = _read_tex_group(block, cursor, "[", "]")
        width, _ = _read_tex_group(block, cursor, "{", "}")
        match = re.fullmatch(
            r"\s*(?:(\d+(?:\.\d+)?)\s*)?\\(?:textwidth|linewidth|columnwidth)\s*",
            width or "",
        )
        blocks.append({
            "block": block,
            "ratio": float(match.group(1) or 1) if match else None,
        })
    return blocks


def _subfigure_caption_records(figure_record):
    """Number the source's non-starred panel captions within each figure."""
    records = []
    figure = figure_record["float"]
    for start, end in _named_environment_ranges(figure, {"subfigure"}):
        block = figure[start:end]
        for match in _CAPTION_PATTERN.finditer(block):
            if match.group(1) or _is_comment_offset(block, match.start()):
                continue
            cursor = match.end()
            while cursor < len(block) and block[cursor].isspace():
                cursor += 1
            if cursor < len(block) and block[cursor] == "[":
                _, cursor = _read_tex_group(block, cursor, "[", "]")
            caption, _ = _read_tex_group(block, cursor, "{", "}")
            if caption is not None:
                records.append({
                    "caption": caption.strip(),
                    "prefix": f"({chr(ord('a') + len(records))})",
                })
    return records


def _preserve_wide_subfigure_tables(source_text, output_path):
    """Keep wide LaTeX subfigures stacked and their nested tables readable."""
    document = Document(output_path)
    figure_tables = [table._tbl for table in document.tables
                     if (style := table._tbl.find(
                         qn("w:tblPr")).find(qn("w:tblStyle"))) is not None
                     and style.get(qn("w:val")) == "FigureTable"]
    source_figures = [record for record in _figure_caption_records(source_text)
                      if _subfigure_blocks(record)]
    if not figure_tables or not source_figures:
        return

    section = document.sections[0]
    text_width = (section.page_width - section.left_margin - section.right_margin) / 914400
    changed = False
    for record, table_element in zip(source_figures, figure_tables):
        subfigures = _subfigure_blocks(record)
        rows = table_element.findall(qn("w:tr"))
        if (len(subfigures) < 2
                or any(item["ratio"] is None or item["ratio"] < 0.7
                       for item in subfigures)
                or len(rows) != 1):
            continue
        cells = rows[0].findall(qn("w:tc"))
        cells = [cell for cell in cells if (
            any((node.text or "").strip() for node in cell.iter(qn("w:t")))
            or cell.find(".//" + qn("w:drawing")) is not None
            or cell.find(".//" + qn("m:oMath")) is not None
            or cell.find(".//" + qn("w:tbl")) is not None
        )]
        if len(cells) != len(subfigures):
            continue

        grid = table_element.find(qn("w:tblGrid"))
        if grid is None:
            continue
        for column in list(grid):
            grid.remove(column)
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(round(text_width * 1440)))
        grid.append(column)
        table_element.remove(rows[0])
        for cell in cells:
            row = OxmlElement("w:tr")
            row.append(deepcopy(cell))
            table_element.append(row)
        _set_table_column_widths(table_element, [text_width])
        changed = True

        for subfigure, row in zip(subfigures, table_element.findall(qn("w:tr"))):
            cell = row.find(qn("w:tc"))
            nested_tables = list(cell.iter(qn("w:tbl")))
            used = set()
            local_width = text_width * subfigure["ratio"]
            for table_width, column_spec in _table_column_specs(
                    {"float": subfigure["block"]}):
                if table_width is None:
                    continue
                widths = _column_widths_inches(
                    table_width, column_spec, local_width)
                if not widths:
                    continue
                target = next((item for item in nested_tables
                               if id(item) not in used
                               and len(item.findall(
                                   f"{qn('w:tblGrid')}/{qn('w:gridCol')}")) == len(widths)),
                              None)
                if target is None:
                    continue
                used.add(id(target))
                ancestor = target.getparent()
                while ancestor is not None and ancestor is not cell:
                    if ancestor.tag == qn("w:tbl"):
                        ancestor_grid = ancestor.find(qn("w:tblGrid"))
                        ancestor_columns = (
                            ancestor_grid.findall(qn("w:gridCol"))
                            if ancestor_grid is not None else [])
                        if len(ancestor_columns) == 1:
                            current_width = int(ancestor_columns[0].get(qn("w:w"), "0"))
                            target_width = round(sum(widths) * 1440)
                            if target_width > current_width:
                                _set_table_column_widths(
                                    ancestor, [sum(widths)])
                    ancestor = ancestor.getparent()
                changed = _set_table_column_widths(target, widths) or changed

    if changed:
        document.save(output_path)


def _plain_caption_text(caption):
    try:
        text = pypandoc.convert_text(
            caption, to="plain", format="latex", extra_args=["--wrap=none"])
    except RuntimeError:
        text = re.sub(r"\\[A-Za-z]+\*?", "", caption)
    return " ".join(text.split())


def _normalized_text(text):
    return " ".join(re.findall(r"[a-z0-9]+", text.lower()))


def _prepend_caption_label(paragraph, label):
    """Insert a bold figure label without flattening existing caption runs."""
    if (re.match(r"^\s*" + re.escape(label) + r"(?:\s|$)", paragraph.text)
            or (label.startswith("Figure ")
                and re.match(r"^\s*Figure\s+\d+\.", paragraph.text, re.I))):
        return False

    for run in paragraph.runs:
        if not run.text:
            continue
        trimmed = run.text.lstrip()
        if trimmed:
            if trimmed != run.text:
                run.text = trimmed
            break
        run._element.getparent().remove(run._element)

    runs = []
    label_run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    properties.append(OxmlElement("w:b"))
    label_run.append(properties)
    label_text = OxmlElement("w:t")
    label_text.text = label
    label_run.append(label_text)
    runs.append(label_run)

    space_run = OxmlElement("w:r")
    space_text = OxmlElement("w:t")
    space_text.set(qn("xml:space"), "preserve")
    space_text.text = " "
    space_run.append(space_text)
    runs.append(space_run)

    first_content_index = 1 if paragraph._p.pPr is not None else 0
    for offset, run in enumerate(runs):
        paragraph._p.insert(first_content_index + offset, run)
    return True


def _match_caption_paragraphs(records, paragraphs):
    if len(records) == len(paragraphs):
        return list(zip(records, paragraphs))
    matches = []
    next_paragraph = 0
    for record in records:
        expected = _normalized_text(_plain_caption_text(record["caption"]))
        if not expected:
            continue
        candidates = [
            (SequenceMatcher(
                None, expected, _normalized_text(paragraph.text)).ratio(),
             index, paragraph)
            for index, paragraph in enumerate(
                paragraphs[next_paragraph:], next_paragraph)
        ]
        if not candidates:
            break
        similarity, index, paragraph = max(candidates, key=lambda item: item[0])
        if similarity < 0.45:
            continue
        matches.append((record, paragraph))
        next_paragraph = index + 1
    return matches


def _restore_figure_caption_labels(source_text, output_path):
    """Restore figure and panel labels while retaining caption runs and math."""
    records = _figure_caption_records(source_text)
    if not records:
        return

    document = Document(output_path)
    paragraphs = [paragraph for paragraph in document.paragraphs
                  if paragraph.style.name == "Image Caption"]
    matches = _match_caption_paragraphs(records, paragraphs)
    panel_records = [panel for record in records
                     for panel in _subfigure_caption_records(record)]
    panel_paragraphs = [
        Paragraph(element, document)
        for element in document.element.body.iter(qn("w:p"))
        if element.getparent() is not document.element.body
        and Paragraph(element, document).style.name == "Image Caption"
    ]
    matches.extend(_match_caption_paragraphs(panel_records, panel_paragraphs))

    changed = False
    for record, paragraph in matches:
        changed = _prepend_caption_label(paragraph, record["prefix"]) or changed
    if changed:
        document.save(output_path)


def _restore_appendix_headings(document, source_text):
    records, _ = _appendix_heading_records(source_text)
    if not records:
        return False

    headings = [paragraph for paragraph in document.paragraphs
                if paragraph.style.name in ("Heading 1", "Heading 2", "Heading 3")]
    next_heading = 0
    changed = False
    for record in records:
        expected = _normalized_text(_plain_caption_text(record["title"]))
        if not expected:
            continue
        style_name = f"Heading {record['level']}"
        candidates = [
            (SequenceMatcher(
                None, expected, _normalized_text(paragraph.text)).ratio(),
             index, paragraph)
            for index, paragraph in enumerate(headings[next_heading:], next_heading)
            if paragraph.style.name == style_name
        ]
        if not candidates:
            continue
        similarity, index, paragraph = max(candidates, key=lambda item: item[0])
        if similarity < 0.45:
            continue
        next_heading = index + 1
        prefix = record["prefix"]
        if not prefix:
            continue
        for run_index, run in enumerate(paragraph.runs):
            if not re.fullmatch(r"\d+(?:\.\d+)*", run.text.strip()):
                continue
            if run.text != prefix:
                run.text = prefix
                changed = True
            for following in paragraph.runs[run_index + 1:]:
                if following.text == "\t":
                    following.text = " "
                    changed = True
                    break
                if following.text.strip():
                    break
            break
    return changed


def _hyperlink_text(hyperlink):
    return "".join(node.text or "" for node in hyperlink.iter(qn("w:t")))


def _set_hyperlink_text(hyperlink, text):
    nodes = list(hyperlink.iter(qn("w:t")))
    if not nodes:
        return False
    if _hyperlink_text(hyperlink) == text:
        return False
    nodes[0].text = text
    for node in nodes[1:]:
        node.text = ""
    return True


def _restore_cross_reference_numbers(source_text, output_path):
    """Correct Pandoc reference numbers for floats and LaTeX appendix labels."""
    document = Document(output_path)
    hyperlinks = list(document.element.iter(qn("w:hyperlink")))
    if not hyperlinks:
        return

    resolved = {}
    for hyperlink in hyperlinks:
        anchor = hyperlink.get(qn("w:anchor"))
        value = _hyperlink_text(hyperlink)
        if anchor and value and not re.fullmatch(r"\[[^\]]+\]", value):
            resolved.setdefault(anchor, value)

    _, appendix_numbers = _appendix_heading_records(source_text)
    label_numbers = dict(appendix_numbers)
    label_numbers.update(_figure_reference_numbers(source_text))
    label_numbers.update(_table_reference_numbers(source_text))
    for alias, target in _section_label_aliases(source_text).items():
        if alias in label_numbers:
            continue
        value = label_numbers.get(target, resolved.get(target))
        if value and not re.fullmatch(r"\[[^\]]+\]", value):
            label_numbers[alias] = value

    changed = False
    for hyperlink in hyperlinks:
        anchor = hyperlink.get(qn("w:anchor"))
        value = label_numbers.get(anchor)
        if value is not None:
            changed = _set_hyperlink_text(hyperlink, value) or changed
    changed = _restore_appendix_headings(document, source_text) or changed
    if changed:
        document.save(output_path)


def _table_cells(table):
    return {
        _normalized_text(cell.text)
        for row in table.rows for cell in row.cells
        if len(_normalized_text(cell.text)) >= 4
    }


def _table_matches(candidate, existing):
    if (len(candidate.rows), len(candidate.columns)) != (
            len(existing.rows), len(existing.columns)):
        return False
    if existing._tbl.findall(".//" + qn("w:drawing")):
        return False
    candidate_cells = _table_cells(candidate)
    existing_cells = _table_cells(existing)
    if not candidate_cells or not existing_cells:
        return False
    shared = len(candidate_cells & existing_cells)
    # ponytail: overlap tolerates merged-cell differences; source-position mapping
    # would be needed if Pandoc starts changing table dimensions between passes.
    return (shared / len(candidate_cells) >= 0.35
            and shared / len(existing_cells) >= 0.55)


def _insert_caption_and_table(document, number, caption_text, table_element=None):
    body = document.element.body
    next_caption = None
    pattern = re.compile(r"^Table\s+(\d+)\.")
    for paragraph in document.paragraphs:
        match = pattern.match(paragraph.text.strip())
        if match and int(match.group(1)) > number:
            next_caption = paragraph._p
            break

    if next_caption is None:
        section_properties = body.sectPr
        index = (list(body).index(section_properties)
                 if section_properties is not None else len(body))
    else:
        index = list(body).index(next_caption)

    paragraph_element = OxmlElement("w:p")
    paragraph = Paragraph(paragraph_element, document._body)
    paragraph.style = _ensure_table_caption_style(document)
    label = paragraph.add_run(f"Table {number}.")
    label.bold = True
    if caption_text:
        paragraph.add_run(" " + caption_text)
    body.insert(index, paragraph_element)
    inserted_table = None
    if table_element is not None:
        inserted_table = deepcopy(table_element)
        body.insert(index + 1, inserted_table)
    return paragraph_element, inserted_table


def _insert_table_after_element(document, anchor, table_element):
    body = document.element.body
    index = list(body).index(anchor) + 1
    inserted = deepcopy(table_element)
    body.insert(index, inserted)
    return inserted


def _restore_table_captions_and_missing_tables(
        source_text, records, output_path, source_dir,
        reference_docx_path, temp_dir, on_warning=None):
    """Restore captions and recover native tables Pandoc silently skipped."""
    document = Document(output_path)
    changed = False
    table_caption_style = _ensure_table_caption_style(document)
    for record in records:
        if record["starred"]:
            continue
        number = record["number"]
        marker = f"Table {number}."
        matching = next((paragraph for paragraph in document.paragraphs
                         if paragraph.text.strip().startswith(marker)), None)
        if matching is not None:
            if matching.style.name != "Table Caption":
                changed = True
            matching.style = table_caption_style
            if not matching.text[len(marker):].strip():
                caption_text = _plain_caption_text(record["caption"])
                for run in list(matching.runs):
                    matching._p.remove(run._r)
                label = matching.add_run(marker)
                label.bold = True
                if caption_text:
                    matching.add_run(" " + caption_text)
                changed = True

        missing_tables = []
        for table_index, table_source in enumerate(
                _standalone_table_sources(record, source_text), start=1):
            candidate_source = os.path.join(
                temp_dir, f"fallback-table-{number}-{table_index}.tex")
            candidate_docx = os.path.join(
                temp_dir, f"fallback-table-{number}-{table_index}.docx")
            candidate_log = os.path.join(
                temp_dir, f"fallback-table-{number}-{table_index}.json")
            with open(candidate_source, "w", encoding="utf-8", newline="") as source_file:
                source_file.write(table_source)
            pypandoc.convert_file(
                candidate_source,
                to="docx",
                format="latex",
                outputfile=candidate_docx,
                extra_args=[
                    f"--resource-path={source_dir}",
                    f"--reference-doc={reference_docx_path}",
                    f"--log={candidate_log}",
                    "--number-sections",
                ],
                cworkdir=source_dir,
            )
            if on_warning:
                for message in _pandoc_warning_messages(candidate_log):
                    on_warning(message)
            candidate = Document(candidate_docx)
            if candidate.tables:
                table = max(
                    candidate.tables,
                    key=lambda item: len(item.rows) * len(item.columns))
                if not any(_table_matches(table, existing)
                           for existing in document.tables):
                    missing_tables.append(table._tbl)

        if matching is None:
            caption_text = _plain_caption_text(record["caption"])
            paragraph_element, anchor = _insert_caption_and_table(
                document, number, caption_text,
                missing_tables.pop(0) if missing_tables else None)
            if anchor is None:
                anchor = paragraph_element
            for table_element in missing_tables:
                anchor = _insert_table_after_element(
                    document, anchor, table_element)
            changed = True
        elif missing_tables:
            anchor = matching._p
            for table_element in missing_tables:
                anchor = _insert_table_after_element(
                    document, anchor, table_element)
            changed = True

    if changed:
        document.save(output_path)


def _preserve_wide_subfigure_image_widths(source_text, output_path):
    image_count, width_overrides = _wide_subfigure_linewidth_images(source_text)
    if not width_overrides:
        return

    document = Document(output_path)
    if len(document.inline_shapes) != image_count:
        return
    section = document.sections[0]
    text_width = section.page_width - section.left_margin - section.right_margin
    for image_index, ratio in width_overrides:
        image = document.inline_shapes[image_index]
        old_width, old_height = int(image.width), int(image.height)
        new_width = Inches(text_width / 914400 * ratio)
        image.width = new_width
        image.height = round(old_height * int(new_width) / old_width)
    document.save(output_path)


def _bibliography_files(source_path):
    """Return bibliography files declared by the source that exist beside it."""
    source_dir = os.path.dirname(os.path.abspath(source_path))
    with open(source_path, encoding="utf-8", errors="replace") as source_file:
        text = _strip_tex_comments(source_file.read())

    files = []
    for value in _BIBLIOGRAPHY_PATTERN.findall(text):
        for entry in value.split(","):
            entry = entry.strip()
            if not entry:
                continue
            if not os.path.splitext(entry)[1]:
                entry += ".bib"
            path = os.path.abspath(os.path.join(source_dir, entry))
            if os.path.isfile(path) and path not in files:
                files.append(path)
    return files


def _bibtex_entry_titles(source_path):
    """Read entry titles for a content audit without relying on Word bookmarks."""
    if not os.path.isfile(source_path):
        return {}
    with open(source_path, encoding="utf-8", errors="replace") as source_file:
        source = source_file.read()
    entry_pattern = re.compile(
        r"(?is)@\w+\s*[({]\s*([^,\s]+)\s*,(.*?)(?=\n\s*@\w+\s*[({]|\Z)")
    title_pattern = re.compile(
        r"(?is)\btitle\s*=\s*(\{(?:[^{}]|\{[^{}]*\})*\}|\"(?:\\.|[^\"])*\")")
    titles = {}
    for entry in entry_pattern.finditer(source):
        match = title_pattern.search(entry.group(2))
        if not match:
            continue
        title = match.group(1).strip()
        if title.startswith(("{", '"')) and title.endswith(("}", '"')):
            title = title[1:-1]
        titles[entry.group(1)] = title
    return titles


def _declared_bibliography_paths(source_text, source_dir):
    source = _strip_tex_comments(source_text)
    paths = []
    for declaration in _BIBLIOGRAPHY_PATTERN.finditer(source):
        for value in declaration.group(1).split(","):
            value = value.strip()
            if not value:
                continue
            if not os.path.splitext(value)[1]:
                value += ".bib"
            path = os.path.abspath(os.path.join(source_dir, value))
            if path not in paths:
                paths.append(path)
    return paths


def _source_image_paths(source_text):
    source = _strip_tex_comments(source_text)
    images = []
    for match in re.finditer(r"\\includegraphics\*?", source):
        cursor = match.end()
        while cursor < len(source) and source[cursor].isspace():
            cursor += 1
        if cursor < len(source) and source[cursor] == "[":
            _, cursor = _read_tex_group(source, cursor, "[", "]")
        image, _ = _read_tex_group(source, cursor, "{", "}")
        if image:
            images.append(image.strip())
    return images


def _source_resource_warnings(source_text, source_dir):
    warnings = []
    for image in _source_image_paths(source_text):
        if "\\" in image:
            continue
        image_path = os.path.abspath(os.path.join(source_dir, image))
        candidates = ([image_path] if os.path.splitext(image_path)[1] else
                      [image_path + extension for extension in
                       (".pdf", ".png", ".jpg", ".jpeg", ".svg", ".eps", ".tif", ".tiff")])
        if not any(os.path.isfile(path) for path in candidates):
            warnings.append(f"Figure resource not found: {image}")

    for path in _declared_bibliography_paths(source_text, source_dir):
        if not os.path.isfile(path):
            warnings.append(f"Bibliography resource not found: {path}")
    return warnings


def _source_citation_keys(source_text):
    source = _strip_tex_comments(source_text)
    pattern = re.compile(
        r"\\(?:cite[a-zA-Z]*|[a-zA-Z]*cite[a-zA-Z]*)\*?"
        r"(?:\s*\[[^]]*\])*\s*\{([^{}]+)\}")
    return {
        key.strip()
        for match in pattern.finditer(source)
        for key in match.group(1).split(",") if key.strip()
    }


def _source_cross_reference_keys(source_text):
    source = _strip_tex_comments(source_text)
    pattern = re.compile(
        r"\\(?:ref|eqref|autoref|pageref|cref|Cref)\*?"
        r"(?:\s*\[[^]]*\])*\s*\{([^{}]+)\}")
    return {
        key.strip()
        for match in pattern.finditer(source)
        for key in match.group(1).split(",") if key.strip()
    }


def _pandoc_warning_messages(log_path, source_path=None, temporary_source=None):
    try:
        with open(log_path, encoding="utf-8") as log_file:
            messages = json.load(log_file)
    except (OSError, ValueError, TypeError):
        return []
    warnings = []
    for item in messages if isinstance(messages, list) else []:
        if str(item.get("verbosity", "")).upper() not in ("WARNING", "ERROR"):
            continue
        message = str(item.get("pretty") or item.get("message") or item.get("type") or "Pandoc warning")
        if temporary_source:
            message = message.replace(temporary_source, source_path or temporary_source)
        warnings.append(message)
    return warnings


def _multi_panel_figure_records(source_text):
    """Figures whose subfigure panels should appear as one picture, as in the PDF."""
    return [record for record in _figure_caption_records(source_text)
            if _named_environment_ranges(record["float"], {"subfigure"})]


_MACRO_DEFINITION_PATTERN = re.compile(
    r"\\(?:newcommand|renewcommand|providecommand|newcolumntype)\*?(?![A-Za-z])")


def _body_macro_definitions(source_text):
    """Return macro definitions made after \\begin{document}, e.g. table cell helpers."""
    parts = source_text.split(r"\begin{document}", 1)
    if len(parts) < 2:
        return ""
    body = parts[1]
    definitions = []
    for match in _MACRO_DEFINITION_PATTERN.finditer(body):
        if _is_comment_offset(body, match.start()):
            continue
        cursor = match.end()
        groups = 0
        while cursor < len(body) and groups < 2:
            while cursor < len(body) and body[cursor].isspace():
                cursor += 1
            if cursor >= len(body):
                break
            if body[cursor] == "[":
                _, cursor = _read_tex_group(body, cursor, "[", "]")
            elif body[cursor] == "{":
                _, cursor = _read_tex_group(body, cursor, "{", "}")
                groups += 1
            elif body[cursor] == "\\" and groups == 0:
                end = cursor + 1
                while end < len(body) and body[end].isalpha():
                    end += 1
                cursor, groups = end, 1
            else:
                break
        definition = body[match.start():cursor]
        if groups == 2 and not re.search(r"\\(?:array|baseline)stretch\b", definition):
            definitions.append(definition)
    return "\n".join(definitions)


def _standalone_figure_source(source_text, figure_block):
    """Return a one-page LaTeX document showing a figure's panels without its caption."""
    return _standalone_float_source(source_text, figure_block, "figure")


def _standalone_float_source(source_text, float_block, float_type):
    """Return a one-page LaTeX document showing a float's body without its caption."""
    preamble = source_text.split(r"\begin{document}", 1)[0]
    inner = re.sub(r"^\s*\\begin\s*\{(?:figure|table)\*?\}\s*(\[[^\]]*\])?", "",
                   float_block)
    inner = re.sub(r"\\end\s*\{(?:figure|table)\*?\}\s*$", "", inner)
    subfigures = _named_environment_ranges(inner, {"subfigure"})
    removals = []
    for match in _CAPTION_PATTERN.finditer(inner):
        if _is_comment_offset(inner, match.start()) or any(
                left < match.start() < right for left, right in subfigures):
            continue
        cursor = match.end()
        while cursor < len(inner) and inner[cursor].isspace():
            cursor += 1
        if cursor < len(inner) and inner[cursor] == "[":
            _, cursor = _read_tex_group(inner, cursor, "[", "]")
        caption, cursor = _read_tex_group(inner, cursor, "{", "}")
        if caption is not None:
            removals.append((match.start(), cursor))
    for start, end in reversed(removals):
        inner = inner[:start] + inner[end:]
    return (preamble + "\\begin{document}\n" + _body_macro_definitions(source_text)
            + "\n\\pagestyle{empty}\n"
            "\\noindent\\begin{minipage}{\\textwidth}\n"
            + (f"\\captionsetup{{type={float_type}}}\n"
               if re.search(r"\\usepackage(?:\[[^\]]*\])?\{[^}]*\b(?:sub)?caption\b", preamble)
               else "")
            + inner.strip() + "\n"
            "\\end{minipage}\n\\end{document}\n")


def _render_figure_png(tex_source, work_dir, source_dir, jobname, dpi=300):
    """Compile a standalone figure with pdflatex and crop it to its content."""
    try:
        import pymupdf as fitz
    except ImportError:  # older PyMuPDF releases only provide the fitz name
        import fitz

    tex_path = os.path.join(work_dir, jobname + ".tex")
    with open(tex_path, "w", encoding="utf-8") as tex_file:
        tex_file.write(tex_source)
    subprocess.run(
        ["pdflatex", "-interaction=nonstopmode", "-halt-on-error",
         f"-output-directory={work_dir}", tex_path],
        cwd=source_dir, check=True, stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL, timeout=180)
    with fitz.open(os.path.join(work_dir, jobname + ".pdf")) as pdf:
        page = pdf[0]
        rects = [fitz.Rect(block["bbox"]) for block in page.get_text("dict")["blocks"]]
        rects += [drawing["rect"] for drawing in page.get_drawings()]
        rects += [fitz.Rect(image["bbox"]) for image in page.get_image_info()]
        # Table rules are often drawn as zero-height lines, which PyMuPDF calls
        # "empty"; keep them so \\toprule and \\bottomrule stay in the crop.
        rects = [rect for rect in rects
                 if not rect.is_infinite and (rect.width > 0 or rect.height > 0)]
        if not rects:
            raise RuntimeError("the compiled figure page is empty")
        # Union by coordinates: PyMuPDF's rect union skips zero-height rules.
        clip = fitz.Rect(min(rect.x0 for rect in rects), min(rect.y0 for rect in rects),
                         max(rect.x1 for rect in rects), max(rect.y1 for rect in rects))
        clip = (clip + (-3, -3, 3, 3)) & page.rect
        png_path = os.path.join(work_dir, jobname + ".png")
        page.get_pixmap(dpi=dpi, clip=clip).save(png_path)
    return png_path, clip.width / 72


def _insert_multi_panel_figures(source_text, source_dir, output_path, work_dir,
                                on_warning=None):
    """Replace subfigure grids with one LaTeX-rendered picture per figure.

    Returns the figure blocks that were replaced, so the content audit can
    count each of them as a single image.
    """
    records = _multi_panel_figure_records(source_text)
    if not records:
        return []
    document = Document(output_path)
    section = document.sections[0]
    text_width = (section.page_width - section.left_margin - section.right_margin) / 914400
    skipped = {qn("w:bookmarkStart"), qn("w:bookmarkEnd")}
    replaced = []
    for record in records:
        caption = next((paragraph for paragraph in document.paragraphs
                        if paragraph.text.strip().startswith(record["prefix"])), None)
        if caption is None:
            continue
        panels = []
        sibling = caption._element.getprevious()
        while sibling is not None:
            if sibling.tag in skipped:
                sibling = sibling.getprevious()
                continue
            has_drawing = next(sibling.iter(qn("w:drawing")), None) is not None
            if sibling.tag == qn("w:tbl") or (sibling.tag == qn("w:p") and has_drawing):
                panels.append(sibling)
                sibling = sibling.getprevious()
                continue
            break
        if not panels:
            continue
        try:
            png_path, width = _render_figure_png(
                _standalone_figure_source(source_text, record["float"]),
                work_dir, source_dir, f"figure{record['number']}")
        except Exception as exc:
            if on_warning:
                on_warning(f"Figure {record['number']}: panels kept separate "
                           f"because the figure could not be rendered ({exc}).")
            continue
        picture = document.add_paragraph()
        picture.alignment = WD_ALIGN_PARAGRAPH.CENTER
        picture.paragraph_format.keep_with_next = True
        picture.add_run().add_picture(png_path, width=Inches(min(width, text_width)))
        caption._element.addprevious(picture._element)
        for panel in panels:
            panel.getparent().remove(panel)
        replaced.append(record["float"])
    document.save(output_path)
    return replaced


def _insert_table_images(source_text, source_dir, output_path, work_dir, on_warning=None):
    """Replace each numbered table with a LaTeX-rendered picture of the table body.

    The caption stays as editable Word text above the picture. Returns the
    table float blocks that were replaced.
    """
    records = [record for record in _table_float_records(source_text)
               if not record["starred"]]
    if not records:
        return []
    document = Document(output_path)
    section = document.sections[0]
    text_width = (section.page_width - section.left_margin - section.right_margin) / 914400
    skipped = {qn("w:bookmarkStart"), qn("w:bookmarkEnd")}
    replaced = []
    for record in records:
        marker = f"Table {record['number']}."
        caption = next((paragraph for paragraph in document.paragraphs
                        if paragraph.text.strip().startswith(marker)), None)
        if caption is None:
            continue
        float_text = _normalized_text(_plain_caption_text(record["float"]))
        parts = []
        node = caption._element.getnext()
        while node is not None:
            text = "".join(item.text or "" for item in node.iter(qn("w:t"))).strip()
            if node.tag in skipped or node.tag == qn("w:tbl"):
                parts.append(node)
            elif node.tag == qn("w:p") and next(node.iter(qn("w:drawing")), None) is None and (
                    not text or _normalized_text(text)[:20] in float_text):
                # Empty spacers and panel headings such as "A. Species ..." belong
                # to the float; ordinary body text does not appear in its source.
                parts.append(node)
            else:
                break
            node = node.getnext()
        while parts and parts[-1].tag != qn("w:tbl"):
            parts.pop()
        if not any(part.tag == qn("w:tbl") for part in parts):
            continue
        table = parts[0]
        try:
            png_path, width = _render_figure_png(
                _standalone_float_source(source_text, record["float"], "table"),
                work_dir, source_dir, f"table{record['number']}")
        except Exception as exc:
            if on_warning:
                on_warning(f"Table {record['number']}: kept as an editable table "
                           f"because it could not be rendered ({exc}).")
            continue
        picture = document.add_paragraph()
        picture.alignment = WD_ALIGN_PARAGRAPH.CENTER
        picture.add_run().add_picture(png_path, width=Inches(min(width, text_width)))
        table.addprevious(picture._element)
        for part in parts:
            part.getparent().remove(part)
        replaced.append(record["float"])
    document.save(output_path)
    return replaced


_DISPLAY_MATH_ENVIRONMENTS = {
    "equation", "equation*", "align", "align*", "gather", "gather*",
    "multline", "multline*", "eqnarray", "eqnarray*", "displaymath",
}


def _display_math_blocks(source_text):
    """Return display-math source blocks in document order, outside floats."""
    body = source_text.split(r"\begin{document}", 1)[-1]
    floats = _named_environment_ranges(body, {"figure", "figure*", "table", "table*"})
    blocks = [(start, end) for start, end in
              _named_environment_ranges(body, _DISPLAY_MATH_ENVIRONMENTS)]
    environments = list(blocks)
    for match in re.finditer(r"(?<!\\)\\\[(.*?)(?<!\\)\\\]", body, re.S):
        if not _is_comment_offset(body, match.start()) and not any(
                left <= match.start() < right for left, right in environments):
            blocks.append((match.start(), match.end()))
    blocks = [(start, end) for start, end in blocks
              if not any(left <= start < right for left, right in floats)]
    return [body[start:end] for start, end in sorted(blocks)]


def _insert_equation_images(source_text, source_dir, output_path, work_dir, on_warning=None):
    """Replace display equations with LaTeX-rendered pictures, in document order."""
    blocks = _display_math_blocks(source_text)
    if not blocks:
        return 0
    document = Document(output_path)
    paragraphs = [element for element in document.element.body
                  if element.tag == qn("w:p")
                  and next(element.iter(qn("m:oMathPara")), None) is not None]
    if len(paragraphs) != len(blocks):
        if on_warning:
            on_warning(f"Equations kept editable: source has {len(blocks)} display "
                       f"equations but the DOCX has {len(paragraphs)}.")
        return 0
    preamble = source_text.split(r"\begin{document}", 1)[0]
    macros = _body_macro_definitions(source_text)
    section = document.sections[0]
    text_width = (section.page_width - section.left_margin - section.right_margin) / 914400
    replaced = 0
    for number, (block, paragraph) in enumerate(zip(blocks, paragraphs), start=1):
        tex_source = (preamble + "\\begin{document}\n" + macros + "\n\\pagestyle{empty}\n"
                      "\\noindent\\begin{minipage}{\\textwidth}\n" + block
                      + "\n\\end{minipage}\n\\end{document}\n")
        try:
            png_path, width = _render_figure_png(
                tex_source, work_dir, source_dir, f"equation{number}")
        except Exception as exc:
            if on_warning:
                on_warning(f"Equation {number}: kept editable because it could not "
                           f"be rendered ({exc}).")
            continue
        picture = document.add_paragraph()
        picture.alignment = WD_ALIGN_PARAGRAPH.CENTER
        picture.add_run().add_picture(png_path, width=Inches(min(width, text_width)))
        paragraph.addprevious(picture._element)
        paragraph.getparent().remove(paragraph)
        replaced += 1
    document.save(output_path)
    return replaced


_ACCENT_MARKS = {
    "'": "\u0301", "`": "\u0300", '"': "\u0308", "^": "\u0302", "~": "\u0303",
    "=": "\u0304", ".": "\u0307", "u": "\u0306", "v": "\u030c", "H": "\u030b",
    "c": "\u0327", "k": "\u0328", "r": "\u030a",
}
_TEX_SYMBOLS = {
    r"\i": "i", r"\j": "j", r"\ss": "ß", r"\o": "ø", r"\O": "Ø", r"\ae": "æ",
    r"\AE": "Æ", r"\oe": "œ", r"\l": "ł", r"\L": "Ł", r"\&": "&", r"\%": "%",
    r"\_": "_", r"\#": "#", r"\$": "$",
}


def _tex_accents_to_unicode(text):
    """Turn BibTeX accent commands such as {\'\i} or {\"e} into Unicode letters."""
    import unicodedata

    def accent(match):
        letter = _TEX_SYMBOLS.get("\\" + match.group(3), match.group(3)) if match.group(2) \
            else match.group(3)
        return unicodedata.normalize("NFC", letter + _ACCENT_MARKS[match.group(1)])

    pattern = re.compile(
        r"\{?\\(['`\"^~=.]|[uvHckr](?=[\s{]))\s*\{?(\\)?([A-Za-z])\}?\}?")
    text = pattern.sub(accent, text)
    for command, symbol in sorted(_TEX_SYMBOLS.items(), key=lambda item: -len(item[0])):
        text = re.sub(re.escape(command) + r"(?![A-Za-z])\s?", symbol, text)
    return text


def _natbib_bbl_entries(source_path, source_text):
    """Read natbib author-year entries from the .bbl LaTeX wrote for this source.

    Returns {key: {"short", "year", "body"}} in .bbl order, or None when the
    document does not use natbib or the .bbl is missing or incomplete.
    """
    source = _strip_tex_comments(source_text)
    if not re.search(r"\\usepackage\s*(?:\[[^\]]*\])?\s*\{natbib\}", source):
        return None
    bbl_path = os.path.splitext(source_path)[0] + ".bbl"
    if not os.path.isfile(bbl_path):
        return None
    with open(bbl_path, encoding="utf-8", errors="replace") as bbl_file:
        bbl = bbl_file.read()
    entries = {}
    items = list(re.finditer(r"\\bibitem\s*\[", bbl))
    for index, item in enumerate(items):
        label, cursor = _read_tex_group(bbl, item.end() - 1, "[", "]")
        key, cursor = _read_tex_group(bbl, cursor, "{", "}")
        if label is None or key is None:
            continue
        end = items[index + 1].start() if index + 1 < len(items) else bbl.find(
            r"\end{thebibliography}", cursor)
        label = " ".join(label.split())
        parsed = re.match(r"^(.*?)\((.*?)\)(.*)$", label, re.S)
        if not parsed:
            return None
        year = re.sub(r"\\natexlab\{([^}]*)\}", r"\1", parsed.group(2))
        entries[key.strip()] = {"short": parsed.group(1).strip(), "year": year.strip(),
                                "body": bbl[cursor:end].strip()}
    if not entries or not _source_citation_keys(source_text) <= set(entries):
        return None
    return entries


_NATBIB_CITE_PATTERN = re.compile(
    r"\\(citep|citet|citealp|citealt|citeauthor|citeyear|citeyearpar|cite)(\*?)(?![A-Za-z])")


def _replace_natbib_citations(text, entries, source_text):
    """Write natbib author-year citations as literal text, as the PDF prints them."""
    options = re.search(r"\\usepackage\s*\[([^\]]*)\]\s*\{natbib\}", source_text)
    square = bool(options and "square" in options.group(1))
    left, right = ("[", "]") if square else ("(", ")")
    pieces = []
    cursor = 0
    for match in _NATBIB_CITE_PATTERN.finditer(text):
        if match.start() < cursor or _is_comment_offset(text, match.start()):
            continue
        position = match.end()
        optional = []
        while True:
            while position < len(text) and text[position].isspace():
                position += 1
            if position < len(text) and text[position] == "[":
                value, position = _read_tex_group(text, position, "[", "]")
                optional.append(value or "")
            else:
                break
        while position < len(text) and text[position].isspace():
            position += 1
        keys, position = _read_tex_group(text, position, "{", "}")
        if keys is None:
            continue
        keys = [key.strip() for key in keys.split(",") if key.strip()]
        if not keys or any(key not in entries for key in keys):
            continue
        prenote, postnote = ("", optional[0]) if len(optional) == 1 else (
            tuple(optional[:2]) if len(optional) >= 2 else ("", ""))
        command = match.group(1)
        groups = []
        for key in keys:
            entry = entries[key]
            if groups and groups[-1][0] == entry["short"]:
                groups[-1][1].append(entry["year"])
            else:
                groups.append((entry["short"], [entry["year"]]))
        if command in ("citet", "cite", "citealt"):
            parts = [f"{author} {left}{', '.join(years)}{right}" if command != "citealt"
                     else f"{author} {', '.join(years)}" for author, years in groups]
            rendered = "; ".join(parts)
        elif command == "citeauthor":
            rendered = "; ".join(author for author, _ in groups)
        elif command in ("citeyear", "citeyearpar"):
            rendered = ", ".join(year for _, years in groups for year in years)
            if command == "citeyearpar":
                rendered = left + rendered + right
        else:
            rendered = "; ".join(f"{author}, {', '.join(years)}" for author, years in groups)
            if prenote:
                rendered = prenote + " " + rendered
            if postnote:
                rendered += ", " + postnote
            if command == "citep":
                rendered = left + rendered + right
        pieces.append(text[cursor:match.start()])
        pieces.append(rendered)
        cursor = position
    pieces.append(text[cursor:])
    return "".join(pieces)


def _bbl_entry_runs(body):
    """Split a .bbl entry into (text, italic) runs with LaTeX markup resolved."""
    text = re.sub(r"\\newblock\s*", " ", body)
    text = re.sub(r"\\penalty\d+\s*", "", text)
    text = re.sub(r"\\natexlab\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\(?:url|doi|href)\{([^}]*)\}(?:\{[^}]*\})?", r"\1", text)
    text = _tex_accents_to_unicode(text)
    text = text.replace("---", "\u2014").replace("--", "\u2013").replace("~", " ")
    text = text.replace("``", "\u201c").replace("''", "\u201d")
    runs = []

    def walk(fragment, italic):
        buffer = []
        index = 0
        while index < len(fragment):
            emphasis = re.match(r"\\(emph|textit|textbf|textsc|textrm|mbox)\s*\{",
                                fragment[index:])
            if emphasis:
                runs.append(("".join(buffer), italic))
                buffer = []
                inner, index = _read_tex_group(fragment, index + emphasis.end() - 1, "{", "}")
                walk(inner or "", not italic if emphasis.group(1) in ("emph", "textit")
                     else italic)
                continue
            if fragment.startswith("{\\em", index) or fragment.startswith("{\\it", index):
                runs.append(("".join(buffer), italic))
                buffer = []
                inner, index = _read_tex_group(fragment, index, "{", "}")
                walk(re.sub(r"^\\(?:em|it)\b\s*", "", inner or ""), not italic)
                continue
            character = fragment[index]
            if character in "{}":
                index += 1
                continue
            if character == "\\":
                command = re.match(r"\\[A-Za-z]+\s*", fragment[index:])
                index += command.end() if command else 2
                continue
            buffer.append(character)
            index += 1
        runs.append(("".join(buffer), italic))

    walk(text, False)
    merged = []
    for value, italic in runs:
        value = re.sub(r"\s+", " ", value)
        if not value:
            continue
        if merged and merged[-1][1] == italic:
            merged[-1] = (merged[-1][0] + value, italic)
        else:
            merged.append((value, italic))
    if merged:
        merged[0] = (merged[0][0].lstrip(), merged[0][1])
        merged[-1] = (merged[-1][0].rstrip(), merged[-1][1])
    return merged


def _insert_bbl_references(source_text, output_path, entries):
    """Write the reference list from the .bbl, where \\bibliography sits in the source."""
    document = Document(output_path)
    body = document.element.body
    for paragraph in document.paragraphs:
        if paragraph.style.name == "Bibliography":
            paragraph._element.getparent().remove(paragraph._element)
    try:
        style = _style_by_name(document, "Bibliography")
    except StopIteration:
        style = document.styles.add_style("Bibliography", WD_STYLE_TYPE.PARAGRAPH)
    _set_style_font(style, 12)
    style.paragraph_format.left_indent = Inches(0.25)
    style.paragraph_format.first_line_indent = Inches(-0.25)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(2)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    source = _strip_tex_comments(source_text)
    bibliography = _BIBLIOGRAPHY_PATTERN.search(source)
    appendix = re.search(r"\\appendix\b", source)
    anchor = None
    if bibliography and appendix and bibliography.start() < appendix.start():
        anchor = next((paragraph._element for paragraph in document.paragraphs
                       if paragraph.style.name.startswith("Heading")
                       and paragraph.text.strip().startswith("Appendix")), None)
    if anchor is None:
        anchor = body.find(qn("w:sectPr"))

    def place(paragraph):
        if anchor is not None:
            anchor.addprevious(paragraph._element)

    if not any(paragraph.text.strip() == "References"
               and paragraph.style.name.startswith("Heading")
               for paragraph in document.paragraphs):
        place(document.add_paragraph("References", style="Heading 1"))
    for key, entry in entries.items():
        paragraph = document.add_paragraph(style="Bibliography")
        for value, italic in _bbl_entry_runs(entry["body"]):
            paragraph.add_run(value).italic = italic or None
        place(paragraph)
    document.save(output_path)


def _caption_label_separator(source_text):
    """Return the caption label separator LaTeX prints (":" unless set otherwise)."""
    source = _strip_tex_comments(source_text)
    setting = re.search(r"labelsep\s*=\s*([A-Za-z]+)", source)
    return {"period": ".", "colon": ":", "space": "", "quad": "",
            "newline": ":", "endash": " –"}.get(setting.group(1), ":") \
        if setting else ":"


def _use_latex_caption_labels(source_text, output_path):
    """Print caption labels as LaTeX does, e.g. "Figure 1:" instead of "Figure 1."."""
    separator = _caption_label_separator(source_text)
    if separator == ".":
        return
    document = Document(output_path)
    changed = False
    for paragraph in document.paragraphs:
        if not paragraph.runs:
            continue
        run = paragraph.runs[0]
        if re.fullmatch(r"(Figure|Table) \d+\.", run.text or ""):
            run.text = run.text[:-1] + separator
            run.bold = True
            changed = True
    if changed:
        document.save(output_path)


def _place_references_like_source(source_text, output_path):
    """Move the bibliography to where \\bibliography sits, e.g. before appendices.

    Pandoc always appends citeproc references at the end of the document.
    """
    source = _strip_tex_comments(source_text)
    bibliography = _BIBLIOGRAPHY_PATTERN.search(source)
    appendix = re.search(r"\\appendix\b", source)
    if not bibliography or not appendix or bibliography.start() > appendix.start():
        return
    document = Document(output_path)
    body = document.element.body
    entries = [paragraph for paragraph in document.paragraphs
               if paragraph.style.name == "Bibliography"]
    appendix_heading = next(
        (paragraph for paragraph in document.paragraphs
         if paragraph.style.name.startswith("Heading")
         and paragraph.text.strip().startswith("Appendix")), None)
    if not entries or appendix_heading is None:
        return
    first = entries[0]._element
    last = entries[-1]._element
    while first.getprevious() is not None and first.getprevious().tag == qn("w:bookmarkStart"):
        first = first.getprevious()
    while last.getnext() is not None and last.getnext().tag == qn("w:bookmarkEnd"):
        last = last.getnext()
    block = []
    node = first
    while node is not None:
        block.append(node)
        if node is last:
            break
        node = node.getnext()
    anchor = appendix_heading._element
    if not any(paragraph.text.strip() == "References"
               and paragraph.style.name.startswith("Heading")
               for paragraph in document.paragraphs):
        heading = document.add_paragraph("References", style="Heading 1")
        anchor.addprevious(heading._element)
    for node in block:
        anchor.addprevious(node)
    document.save(output_path)


_LEADING_PUNCTUATION = re.compile(r"^[.,;:!?)\]]+")


def _attach_punctuation_to_links(output_path):
    """Keep punctuation after a citation or cross-reference link on the same line.

    Word may break the line between a hyperlink and the next run, which leaves
    a full stop or comma at the start of the next line.
    """
    document = Document(output_path)
    changed = False
    for link in document.element.body.iter(qn("w:hyperlink")):
        texts = list(link.iter(qn("w:t")))
        following = link.getnext()
        if not texts or following is None or following.tag != qn("w:r"):
            continue
        following_text = following.find(qn("w:t"))
        if following_text is None or not following_text.text:
            continue
        match = _LEADING_PUNCTUATION.match(following_text.text)
        if not match:
            continue
        texts[-1].text = (texts[-1].text or "") + match.group(0)
        texts[-1].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        remainder = following_text.text[match.end():]
        if remainder:
            following_text.text = remainder
        else:
            following.getparent().remove(following)
        changed = True
    if changed:
        document.save(output_path)


def _audit_converted_document(source_text, source_dir, output_path, merged_figures=(),
                              table_images=()):
    """Report concrete source content that did not survive into the DOCX."""
    document = Document(output_path)
    warnings = []
    docx_text = " ".join(node.text or "" for node in document.element.iter(qn("w:t")))
    source_tables = len(_named_environment_ranges(source_text, {"tabular", "tabularx"}))
    source_tables -= sum(len(_named_environment_ranges(block, {"tabular", "tabularx"}))
                         for block in list(merged_figures) + list(table_images)
                         + [_body_macro_definitions(source_text)])
    output_tables = sum(1 for _ in document.element.iter(qn("w:tbl")))
    if output_tables < source_tables:
        warnings.append(
            f"Editable tables: source has {source_tables}; DOCX has {output_tables}. "
            "Review the table panels for missing content.")

    paragraphs = [Paragraph(element, document)
                  for element in document.element.body.iter(qn("w:p"))]
    for record in _table_float_records(source_text):
        if record["starred"]:
            continue
        marker = f"Table {record['number']}."
        matches = [paragraph for paragraph in paragraphs
                   if paragraph.text.strip().startswith(marker)]
        if not matches:
            warnings.append(f"Table {record['number']}: caption or table was not recovered.")
            continue
        caption = _normalized_text(_plain_caption_text(record["caption"]))
        if caption and not any(caption[:min(7, len(caption))]
                               in _normalized_text(paragraph.text) for paragraph in matches):
            warnings.append(
                f"Table {record['number']}: caption text may be incomplete.")
        if any(caption and _normalized_text(paragraph.text).count(caption) > 1
               for paragraph in matches):
            warnings.append(f"Table {record['number']}: caption text is duplicated.")

    source_images = _source_image_paths(source_text)
    expected_images = len(source_images) - sum(
        len(_source_image_paths(block)) - 1 for block in merged_figures)
    output_images = len(document.inline_shapes)
    if output_images < expected_images:
        warnings.append(
            f"Figures: source references {expected_images} images; "
            f"DOCX embeds {output_images}.")

    for record in _figure_caption_records(source_text):
        marker = record["prefix"]
        caption = _normalized_text(_plain_caption_text(record["caption"]))
        matches = [paragraph for paragraph in paragraphs
                   if paragraph.text.strip().startswith(marker)]
        if not matches:
            warnings.append(f"{marker[:-1]}: caption was not recovered.")
        elif caption and not any(
                caption[:min(7, len(caption))] in _normalized_text(paragraph.text)
                for paragraph in matches):
            warnings.append(f"{marker[:-1]}: caption text may be incomplete.")
        if record["float"] in merged_figures:
            continue
        for panel in _subfigure_caption_records(record):
            expected = _normalized_text(_plain_caption_text(panel["caption"]))
            if not any(
                    paragraph.text.strip().startswith(panel["prefix"])
                    and expected[:min(7, len(expected))]
                    in _normalized_text(paragraph.text) for paragraph in paragraphs):
                warnings.append(
                    f"{marker[:-1]} {panel['prefix']}: panel label or caption is missing.")

    label_keys = set(_LABEL_PATTERN.findall(_strip_tex_comments(source_text)))
    for key in sorted(_source_cross_reference_keys(source_text) - label_keys):
        warnings.append(f"Cross-reference '{key}' has no matching source label.")

    bibliography_paths = _declared_bibliography_paths(source_text, source_dir)
    citation_keys = _source_citation_keys(source_text)
    entry_titles = {}
    for path in bibliography_paths:
        if os.path.isfile(path):
            entry_titles.update(_bibtex_entry_titles(path))
    for key in sorted(citation_keys - set(entry_titles)):
        warnings.append(f"Reference '{key}' is cited but missing from the bibliography resources.")

    bibliography_paragraphs = [paragraph.text for paragraph in document.paragraphs
                               if paragraph.style.name == "Bibliography"]
    if len(bibliography_paragraphs) < len(citation_keys):
        warnings.append(
            f"References: {len(citation_keys)} unique citations but only "
            f"{len(bibliography_paragraphs)} bibliography entries in the DOCX.")
    normalized_bibliography = _normalized_text(" ".join(bibliography_paragraphs))
    for key in sorted(citation_keys & set(entry_titles)):
        expected_title = _normalized_text(_plain_caption_text(entry_titles[key]))
        if len(expected_title) >= 8 and expected_title not in normalized_bibliography:
            warnings.append(f"Reference '{key}': its title is missing from the DOCX bibliography.")

    hyperlinks = {
        node.get(qn("w:anchor")): re.sub(r"[.,;:!?)\]]+$", "", _hyperlink_text(node))
        for node in document.element.iter(qn("w:hyperlink"))
        if node.get(qn("w:anchor"))
    }
    expected_numbers = _table_reference_numbers(source_text)
    expected_numbers.update(_figure_reference_numbers(source_text))
    _, appendix_numbers = _appendix_heading_records(source_text)
    expected_numbers.update(appendix_numbers)
    for key in sorted(_source_cross_reference_keys(source_text) & set(expected_numbers)):
        expected = expected_numbers[key]
        actual = hyperlinks.get(key)
        if actual is None:
            warnings.append(f"Cross-reference '{key}': expected Word link to {expected} is missing.")
        elif actual != expected:
            warnings.append(
                f"Cross-reference '{key}': expected {expected}, DOCX shows {actual}.")

    return warnings


def convert_latex_to_docx(source_path, destination_path, on_status=None, on_warning=None,
                         on_progress=None, tables_as_images=False,
                         equations_as_images=False):
    """Convert a .tex source, reporting warnings and completed-stage progress.

    With tables_as_images or equations_as_images, numbered tables or display
    equations are inserted as pictures rendered by pdflatex instead of
    editable Word tables or equations.
    """
    source_path = os.path.abspath(source_path)
    destination_path = os.path.abspath(destination_path)
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f"LaTeX source file not found: {source_path}")
    if not source_path.lower().endswith(".tex"):
        raise ValueError("Choose a LaTeX source file with the .tex extension.")

    warnings = []
    completed_stages = 0
    os.makedirs(os.path.dirname(destination_path), exist_ok=True)

    def report_status(message):
        if on_status:
            on_status(message)

    def finish_stage():
        nonlocal completed_stages
        completed_stages += 1
        if on_progress:
            on_progress(round(completed_stages * 100 / 20))

    def add_warning(message):
        message = str(message).strip()
        if message and message not in warnings:
            warnings.append(message)

    def deliver_warnings():
        if on_warning:
            for message in warnings:
                on_warning(message)

    if on_progress:
        on_progress(0)
    report_status("Preparing LaTeX source...")

    source_dir = os.path.dirname(source_path)
    extra_args = [f"--resource-path={source_dir}"]
    bibliography_files = _bibliography_files(source_path)
    for bibliography_path in bibliography_files:
        extra_args.append(f"--bibliography={bibliography_path}")
    if bibliography_files:
        extra_args.append("--citeproc")

    with open(source_path, encoding="utf-8", errors="replace") as source_file:
        source_text = source_file.read()
    bbl_entries = _natbib_bbl_entries(source_path, source_text)
    if bbl_entries:
        # LaTeX already formatted the citations; reuse its .bbl instead of citeproc.
        extra_args = [argument for argument in extra_args
                      if not argument.startswith("--bibliography=")
                      and argument != "--citeproc"]
    for message in _source_resource_warnings(source_text, source_dir):
        add_warning(message)

    table_records = _table_float_records(source_text)
    cleaned_source = _normalize_table_captions(source_text)
    cleaned_source = _normalize_minipage_tables(cleaned_source)
    cleaned_source = _normalize_table_markup(cleaned_source)
    cleaned_source = _normalize_includegraphics_options(cleaned_source)
    if bbl_entries:
        cleaned_source = _replace_natbib_citations(cleaned_source, bbl_entries, source_text)

    with tempfile.TemporaryDirectory(prefix="pdf2docx-latex-") as temp_dir:
        temp_source_path = os.path.join(temp_dir, os.path.basename(source_path))
        reference_docx_path = os.path.join(temp_dir, "reference.docx")
        main_log_path = os.path.join(temp_dir, "pandoc-log.json")
        with open(temp_source_path, "w", encoding="utf-8", newline="") as temp_source:
            temp_source.write(cleaned_source)
        finish_stage()
        report_status("Creating Word layout...")
        _create_reference_docx(source_text, reference_docx_path)
        finish_stage()

        previous_output_stat = None
        try:
            previous_output_stat = (os.stat(destination_path).st_mtime_ns,
                                    os.stat(destination_path).st_size)
        except OSError:
            pass
        report_status("Converting LaTeX with Pandoc...")
        try:
            pypandoc.convert_file(
                temp_source_path,
                to="docx",
                format="latex",
                outputfile=destination_path,
                extra_args=extra_args + [
                    f"--reference-doc={reference_docx_path}",
                    f"--log={main_log_path}",
                    "--number-sections",
                ],
                cworkdir=source_dir,
            )
        except RuntimeError as exc:
            add_warning("Pandoc: " + str(exc).replace(temp_source_path, source_path))
            for message in _pandoc_warning_messages(
                    main_log_path, source_path, temp_source_path):
                add_warning(message)
            output_is_new = False
            try:
                stat = os.stat(destination_path)
                output_is_new = (previous_output_stat is None
                                 or (stat.st_mtime_ns, stat.st_size) != previous_output_stat)
                Document(destination_path)
            except (OSError, ValueError):
                output_is_new = False
            if not output_is_new:
                deliver_warnings()
                raise RuntimeError(str(exc).replace(temp_source_path, source_path)) from exc
            add_warning(f"A usable partial DOCX was saved to {destination_path}.")
        else:
            for message in _pandoc_warning_messages(
                    main_log_path, source_path, temp_source_path):
                add_warning(message)
        finish_stage()

        merged_figures = []
        table_images = []
        repairs = (
            ("Title block", lambda: _restore_title_block(source_text, destination_path)),
            ("Table recovery", lambda: _restore_table_captions_and_missing_tables(
                source_text, table_records, destination_path, source_dir,
                reference_docx_path, temp_dir, on_warning=add_warning)),
            ("Figure captions", lambda: _restore_figure_caption_labels(
                source_text, destination_path)),
            ("Cross-references", lambda: _restore_cross_reference_numbers(
                source_text, destination_path)),
            ("References", lambda: _insert_bbl_references(
                source_text, destination_path, bbl_entries) if bbl_entries
             else _place_references_like_source(source_text, destination_path)),
            ("Table widths", lambda: _preserve_latex_table_widths(
                source_text, destination_path)),
            ("Wide figure panels", lambda: _preserve_wide_subfigure_tables(
                source_text, destination_path)),
            ("Figure image sizing", lambda: _preserve_wide_subfigure_image_widths(
                source_text, destination_path)),
            ("Multi-panel figures", lambda: merged_figures.extend(
                _insert_multi_panel_figures(source_text, source_dir, destination_path,
                                            temp_dir, on_warning=add_warning))),
            ("Table formatting", lambda: _format_latex_tables(destination_path)),
            ("Table caption flow", lambda: _keep_table_captions_with_tables(
                destination_path)),
            ("Nested table widths", lambda: _constrain_document_nested_tables(
                destination_path)),
            ("Table images", lambda: table_images.extend(
                _insert_table_images(source_text, source_dir, destination_path,
                                     temp_dir, on_warning=add_warning)
                if tables_as_images else [])),
            ("Equation images", lambda: _insert_equation_images(
                source_text, source_dir, destination_path, temp_dir,
                on_warning=add_warning) if equations_as_images else None),
            ("Punctuation after links", lambda: _attach_punctuation_to_links(
                destination_path)),
        )
        for operation, repair in repairs:
            report_status(f"{operation}...")
            try:
                repair()
            except Exception as exc:
                add_warning(f"{operation}: {exc}")
            finish_stage()

        report_status("Checking DOCX content...")
        try:
            for message in _audit_converted_document(
                    source_text, source_dir, destination_path, merged_figures,
                    table_images):
                add_warning(message)
        except Exception as exc:
            add_warning(f"DOCX content audit could not finish: {exc}")
        finish_stage()
        report_status("Finalizing DOCX...")
        try:
            _use_latex_caption_labels(source_text, destination_path)
        except Exception as exc:
            add_warning(f"Caption labels: {exc}")

    deliver_warnings()
    finish_stage()
    return destination_path
