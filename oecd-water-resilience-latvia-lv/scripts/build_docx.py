#!/usr/bin/env python3
"""Build the Latvian OECD report DOCX from the ordered Markdown translations."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
TRANSLATION_DIR = ROOT / "translation"
OUTPUT = ROOT / "output" / "OECD_udens_resursu_noturiba_Latvija.docx"
ORDER = [
    "00-saīsinājumi-un-akronīmi.md",
    "01-udensapgade-un-kanalizacija.md",
    "02-udensapgades-un-kanalizacijas-parvaldiba.md",
    "03-ilgtermina-strategija.md",
]


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, color: str = "D9D9D9", size: str = "6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    cant_split.set(qn("w:val"), "true")
    tr_pr.append(cant_split)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def configure_document(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Aptos"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.08

    for name, size, before, after in [
        ("Title", 22, 0, 14),
        ("Heading 1", 16, 18, 8),
        ("Heading 2", 13, 14, 6),
        ("Heading 3", 11.5, 10, 4),
        ("Heading 4", 10.5, 8, 3),
        ("Heading 5", 10.5, 6, 3),
    ]:
        style = styles[name]
        style.font.name = "Aptos Display" if name in {"Title", "Heading 1", "Heading 2"} else "Aptos"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), style.font.name)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    if "Caption" not in styles:
        caption = styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
    else:
        caption = styles["Caption"]
    caption.font.name = "Aptos"
    caption._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
    caption.font.size = Pt(9.5)
    caption.font.bold = True
    caption.font.color.rgb = RGBColor(0, 0, 0)
    caption.paragraph_format.space_before = Pt(8)
    caption.paragraph_format.space_after = Pt(4)
    caption.paragraph_format.keep_with_next = True

    if "Source Note" not in styles:
        note = styles.add_style("Source Note", WD_STYLE_TYPE.PARAGRAPH)
    else:
        note = styles["Source Note"]
    note.font.name = "Aptos"
    note._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
    note.font.size = Pt(9)
    note.font.italic = True
    note.font.color.rgb = RGBColor(80, 80, 80)
    note.paragraph_format.space_after = Pt(5)
    note.paragraph_format.keep_together = True

    if "Footnote Text" not in styles:
        foot = styles.add_style("Footnote Text", WD_STYLE_TYPE.PARAGRAPH)
    else:
        foot = styles["Footnote Text"]
    foot.font.name = "Aptos"
    foot._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
    foot.font.size = Pt(8.5)
    foot.paragraph_format.left_indent = Inches(0.25)
    foot.paragraph_format.first_line_indent = Inches(-0.25)
    foot.paragraph_format.space_after = Pt(3)

    for sec in doc.sections:
        footer = sec.footer.paragraphs[0]
        add_page_number(footer)
        footer.runs[0].font.size = Pt(8)
        footer.runs[0].font.color.rgb = RGBColor(100, 100, 100)


def add_hyperlink(paragraph, text: str, url: str) -> None:
    """Add a real external hyperlink relationship to a paragraph."""
    relationship_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)

    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1F4E79")
    properties.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.append(underline)
    run.append(properties)

    text_node = OxmlElement("w:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_run_text(paragraph, text: str, bold=False, italic=False, code=False) -> None:
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    if code:
        run.font.name = "Courier New"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Courier New")
        run.font.size = Pt(9)


def add_inline(paragraph, text: str) -> None:
    """Render the small Markdown inline subset used in the translation."""
    pattern = re.compile(
        r"(\[\^[0-9]+\]|\[[^\]]+\]\([^\)]+\)|\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|`[^`]+`|\\\*|\\_)"
    )
    pos = 0
    for match in pattern.finditer(text):
        if match.start() > pos:
            add_run_text(paragraph, text[pos : match.start()])
        token = match.group(0)
        if token.startswith("[^"):
            add_run_text(paragraph, token[2:-1], italic=True)
        elif token.startswith("[") and "](" in token:
            label, url = token[1:].split("](", 1)
            add_hyperlink(paragraph, label, url[:-1])
        elif token.startswith("**") or token.startswith("__"):
            add_run_text(paragraph, token[2:-2], bold=True)
        elif token.startswith("*"):
            add_run_text(paragraph, token[1:-1], italic=True)
        elif token.startswith("`"):
            add_run_text(paragraph, token[1:-1], code=True)
        else:
            add_run_text(paragraph, token.replace("\\", ""))
        pos = match.end()
    if pos < len(text):
        add_run_text(paragraph, text[pos:])


def add_text_paragraph(doc: Document, text: str, style="Normal", left_indent=None) -> None:
    paragraph = doc.add_paragraph(style=style)
    if left_indent is not None:
        paragraph.paragraph_format.left_indent = Inches(left_indent)
    add_inline(paragraph, text)


def is_table_separator(line: str) -> bool:
    cells = [c.strip() for c in line.strip().strip("|").split("|")]
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", c) for c in cells)


def parse_table(lines: list[str], start: int):
    rows = []
    i = start
    while i < len(lines) and lines[i].strip().startswith("|"):
        if not is_table_separator(lines[i]):
            cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
            rows.append(cells)
        i += 1
    return rows, i


def add_table(doc: Document, rows: list[list[str]], compact: bool = False) -> None:
    if not rows:
        return
    cols = max(len(row) for row in rows)
    table = doc.add_table(rows=1, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    header = table.rows[0]
    repeat_table_header(header)
    prevent_row_split(header)
    for ci in range(cols):
        cell = header.cells[ci]
        cell.text = ""
        set_cell_shading(cell, "1F4E79")
        set_cell_borders(cell)
        if compact:
            set_cell_margins(cell, top=35, start=100, bottom=35, end=100)
        else:
            set_cell_margins(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        add_inline(p, rows[0][ci] if ci < len(rows[0]) else "")
        for run in p.runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(8 if compact else 8.5)
    for ri, row in enumerate(rows[1:], start=1):
        cells = table.add_row().cells
        prevent_row_split(table.rows[-1])
        for ci in range(cols):
            cell = cells[ci]
            cell.text = ""
            set_cell_shading(cell, "F2F5F8" if ri % 2 == 0 else "FFFFFF")
            set_cell_borders(cell)
            if compact:
                set_cell_margins(cell, top=35, start=100, bottom=35, end=100)
            else:
                set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            add_inline(p, row[ci] if ci < len(row) else "")
            for run in p.runs:
                run.font.size = Pt(8 if compact else 8.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_list_item(doc: Document, text: str, level: int = 0, ordered: bool = False, number: str | None = None) -> None:
    style = "List Number" if ordered else "List Bullet"
    p = doc.add_paragraph(style=style)
    p.paragraph_format.left_indent = Inches(0.25 + 0.2 * level)
    p.paragraph_format.first_line_indent = Inches(-0.15)
    if ordered and number:
        # Keep the source numbering visible instead of relying on Word's auto-numbering.
        p.style = "Normal"
        p.paragraph_format.left_indent = Inches(0.25 + 0.2 * level)
        p.paragraph_format.first_line_indent = Inches(-0.15)
        add_inline(p, f"{number} ")
    add_inline(p, text)


def add_caption(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="Caption")
    add_inline(p, text.strip("*"))


def add_source_note(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="Source Note")
    add_inline(p, text)


def build_markdown_file(doc: Document, path: Path, page_break: bool) -> None:
    if page_break:
        doc.add_page_break()
    lines = path.read_text(encoding="utf-8").splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("|") and i + 1 < len(lines) and is_table_separator(lines[i + 1]):
            rows, i = parse_table(lines, i)
            add_table(doc, rows, compact=path.name.startswith("00-"))
            continue
        if re.match(r"^#{1,6}\s+", stripped):
            m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
            level = len(m.group(1))
            text = m.group(2)
            style = "Title" if level == 1 and not doc.paragraphs else f"Heading {min(level, 5)}"
            add_text_paragraph(doc, text, style=style)
            i += 1
            continue
        if stripped.startswith("**") and stripped.endswith("**"):
            add_caption(doc, stripped)
            i += 1
            continue
        if re.match(r"^\[\^[0-9]+\]:", stripped):
            m = re.match(r"^\[\^([0-9]+)\]:\s*(.*)$", stripped)
            add_source_note(doc, f"Piezīme {m.group(1)}. {m.group(2)}")
            i += 1
            continue
        if stripped.startswith("- "):
            add_list_item(doc, stripped[2:].strip())
            i += 1
            continue
        ordered = re.match(r"^(\d+[.)])\s+(.*)$", stripped)
        if ordered:
            add_list_item(doc, ordered.group(2), ordered=True, number=ordered.group(1))
            i += 1
            continue
        if stripped.startswith("Avots:") or stripped.startswith("Piezīme:") or stripped.startswith("Piezīmes:"):
            add_source_note(doc, stripped)
            i += 1
            continue
        add_text_paragraph(doc, stripped)
        i += 1


def main() -> None:
    for name in ORDER:
        if not (TRANSLATION_DIR / name).is_file():
            raise FileNotFoundError(TRANSLATION_DIR / name)
    doc = Document()
    configure_document(doc)

    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_inline(title, "OECD Ūdens resursu noturības stiprināšana Latvijā")
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(8)
    add_inline(subtitle, "Pilns tulkojums latviešu valodā")
    date = doc.add_paragraph()
    date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_inline(date, "Projekts apspriešanai ar iesaistītajām pusēm, 2026. gada 4. novembris")
    doc.add_page_break()

    for idx, name in enumerate(ORDER):
        build_markdown_file(doc, TRANSLATION_DIR / name, page_break=idx > 0)

    doc.core_properties.title = "OECD Ūdens resursu noturības stiprināšana Latvijā"
    doc.core_properties.subject = "Pilns OECD ziņojuma tulkojums latviešu valodā"
    doc.core_properties.language = "lv-LV"
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
