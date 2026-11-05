#!/usr/bin/env python3
"""Build the Latvian report by translating the original DOCX in place.

The original OECD document is deliberately used as the template. This keeps
its page geometry, styles, tables, figures, diagrams, headers, footers and
embedded media intact; only the document text is replaced with Latvian text.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "source" / "Draft%20report%20-%20Fostering%20water%20resilience%20in%20Latvia%20-%20For%20stakeholders%204%20Nov%202026.docx"
TRANSLATION_DIR = ROOT / "translation"
OUTPUT = ROOT / "output" / "OECD_udens_resursu_noturiba_Latvija.docx"
ORDER = [
    "00-saīsinājumi-un-akronīmi.md",
    "01-udensapgade-un-kanalizacija.md",
    "02-udensapgades-un-kanalizacijas-parvaldiba.md",
    "03-ilgtermina-strategija.md",
]

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


@dataclass
class Block:
    kind: str
    text: str
    rows: list[list[str]] | None = None
    source_file: str = ""


def is_table_separator(line: str) -> bool:
    cells = [c.strip() for c in line.strip().strip("|").split("|")]
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", c) for c in cells)


def parse_table(lines: list[str], start: int) -> tuple[list[list[str]], int]:
    rows: list[list[str]] = []
    i = start
    while i < len(lines) and lines[i].strip().startswith("|"):
        if not is_table_separator(lines[i]):
            rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
        i += 1
    return rows, i


def parse_markdown() -> list[Block]:
    blocks: list[Block] = []
    for filename in ORDER:
        lines = (TRANSLATION_DIR / filename).read_text(encoding="utf-8").splitlines()
        i = 0
        while i < len(lines):
            stripped = lines[i].strip()
            if not stripped:
                i += 1
                continue
            if stripped.startswith("|") and i + 1 < len(lines) and is_table_separator(lines[i + 1]):
                rows, i = parse_table(lines, i)
                blocks.append(Block("table", "", rows, filename))
                continue
            if re.match(r"^#{1,6}\s+", stripped):
                m = re.match(r"^(#+)\s+(.*)$", stripped)
                assert m
                kind = f"h{len(m.group(1))}"
            elif stripped.startswith("**") and stripped.endswith("**"):
                kind = "caption"
            elif stripped.startswith("- "):
                kind = "bullet"
            elif re.match(r"^\d+[.)]\s+", stripped):
                kind = "ordered"
            elif re.match(r"^\[\^[0-9]+\]:", stripped):
                kind = "footnote"
            else:
                kind = "para"
            blocks.append(Block(kind, stripped, source_file=filename))
            i += 1
    return blocks


def plain_text(text: str) -> str:
    """Remove the small Markdown syntax used by the translation files."""
    text = re.sub(r"^#{1,6}\s+", "", text)
    text = re.sub(r"\[\^[0-9]+\]", "", text)
    text = re.sub(r"\[([^\]]+)\]\([^\)]+\)", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"__([^_]+)__", r"\1", text)
    text = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = text.replace("\\*", "*").replace("\\_", "_")
    return re.sub(r"\s+", " ", text).strip()


def block_text(block: Block) -> str:
    text = block.text
    if block.kind == "bullet":
        text = re.sub(r"^-\s+", "", text)
    elif block.kind == "ordered":
        text = re.sub(r"^\d+[.)]\s+", "", text)
    return text


def text_nodes(element) -> list:
    return [node for node in element.iter() if node.tag == qn("w:t")]


def field_result_nodes(paragraph: Paragraph) -> set:
    """Return text nodes belonging to Word fields, whose cached values must remain."""
    result: set = set()
    depth = 0
    for node in paragraph._p.iter():
        if node.tag == qn("w:fldChar"):
            field_type = node.get(qn("w:fldCharType"))
            if field_type == "begin":
                depth += 1
            elif field_type == "end":
                depth = max(0, depth - 1)
        elif node.tag == qn("w:t") and depth:
            result.add(node)
    return result


def set_paragraph_text(paragraph: Paragraph, text: str) -> None:
    """Replace visible text without deleting drawings, bookmarks or footnotes."""
    text = plain_text(text)
    nodes = text_nodes(paragraph._p)
    if nodes:
        field_nodes = field_result_nodes(paragraph)
        if field_nodes and paragraph.style.name == "Caption":
            # Figure/table captions use STYLEREF and SEQ fields for their
            # numbering. LibreOffice can render those fields as the full
            # chapter title, so replace the fields with the translated,
            # already-numbered caption while retaining the drawing/bookmarks.
            field_runs = []
            for run in paragraph._p.iter(qn("w:r")):
                if any(node in field_nodes for node in run.iter()) or any(
                    node.tag in {qn("w:fldChar"), qn("w:instrText")} for node in run.iter()
                ):
                    field_runs.append(run)
            for run in field_runs:
                run.getparent().remove(run)
            visible_nodes = text_nodes(paragraph._p)
            target = next((node for node in visible_nodes if node.text and node.text.strip()), visible_nodes[0])
            target.text = text
            target.set(XML_SPACE, "preserve")
            for node in visible_nodes:
                if node is target:
                    continue
                node.text = ""
                node.set(XML_SPACE, "preserve")
            return
        target = next((node for node in nodes if node.text and node.text.strip()), nodes[0])
        target.text = text
        target.set(XML_SPACE, "preserve")
        for node in nodes:
            if node is target:
                continue
            node.text = ""
            node.set(XML_SPACE, "preserve")
        return
    if not text:
        return
    run = OxmlElement("w:r")
    text_node = OxmlElement("w:t")
    text_node.text = text
    text_node.set(XML_SPACE, "preserve")
    run.append(text_node)
    paragraph._p.insert(0, run)


def set_cell_text(cell: _Cell, text: str) -> None:
    paragraphs = cell.paragraphs
    if not paragraphs:
        return
    parts = [x.strip() for x in text.split("; ")] if len(paragraphs) > 1 and "; " in text else [text]
    for index, paragraph in enumerate(paragraphs):
        set_paragraph_text(paragraph, parts[index] if index < len(parts) else "")


def paragraph_has_drawing(paragraph: Paragraph) -> bool:
    return any(node.tag == qn("w:drawing") for node in paragraph._p.iter())


def source_kind(paragraph: Paragraph) -> tuple[str, str]:
    style = paragraph.style.name if paragraph.style else ""
    text = paragraph.text.strip()
    if style.startswith("Heading"):
        number = re.search(r"(\d+)$", style)
        level = int(number.group(1)) if number else 1
        return ("heading", "h1" if level == 7 else f"h{level}")
    if style == "Caption":
        if re.match(r"(?:Percentage|% change|20-64)\b", text, re.I):
            return ("normal", "normal")
        return ("caption", "caption")
    if style in {"Caption Subtitle", "Figure", "Source & notes"}:
        return ("normal", "normal")
    if style.startswith("Bulleted") or style.startswith("List Bullet") or style.startswith("Numbered"):
        return ("list", "list")
    return ("normal", "normal")


def block_matches(block: Block, expected: str) -> bool:
    if expected == "caption":
        return block.kind == "caption"
    if expected.startswith("h"):
        return block.kind.startswith("h")
    if expected == "list":
        return block.kind in {"bullet", "ordered", "para"}
    return block.kind in {"para", "bullet", "ordered"}


def find_block(blocks: list[Block], start: int, expected: str, limit: int = 24) -> int | None:
    for index in range(start, min(len(blocks), start + limit)):
        if block_matches(blocks[index], expected):
            return index
    return None


def identify_box_ranges(blocks: list[Block]) -> dict[str, tuple[int, int]]:
    def find(label: str, after: int = -1, kind: str | None = None) -> int:
        label = plain_text(label)
        for i in range(after + 1, len(blocks)):
            block = blocks[i]
            if (kind is None or block.kind == kind) and label.lower() in plain_text(block.text).lower():
                return i
        raise ValueError(f"Translation block not found: {label}")

    def find_caption(label: str) -> int:
        return find(label, kind="caption")

    def next_index(start: int, predicate) -> int:
        for i in range(start + 1, len(blocks)):
            if predicate(blocks[i]):
                return i
        return len(blocks)

    starts = {
        "2.1": find_caption("2.1. ielikums"),
        "2.2": find_caption("2.2. ielikums"),
        "2.3": find_caption("2.3. ielikums"),
        "2.4": find_caption("2.4. ielikums"),
        "2.5": find_caption("2.5. ielikums"),
        "3.1": find_caption("3.1. ielikums"),
        "3.2": find_caption("3.2. ielikums"),
        "3.3": find_caption("3.3. ielikums"),
        "3.4": find_caption("3.4. ielikums"),
        "3.5": find_caption("3.5. ielikums"),
        "3.6": find_caption("3.6. ielikums"),
    }
    result: dict[str, tuple[int, int]] = {}
    for key, start in starts.items():
        if key == "2.2":
            end = find("Papildus šai ūdens resursu", after=start)
        elif key == "2.3":
            end = find("Valsts kontrole ex post", after=start)
        elif key == "2.4":
            end = next(
                i for i in range(start + 1, len(blocks))
                if blocks[i].kind == "h4" and plain_text(blocks[i].text).lower() == "finansēšanas avoti"
            )
        elif key == "2.5":
            end = next_index(start, lambda b: b.kind == "footnote" or b.kind == "h1")
        elif key == "3.1":
            end = find("## 1. rekomendācija", after=start)
        elif key == "3.2":
            end = find("#### 2. Stiprināt investīciju plānošanu", after=start)
        elif key == "3.3":
            end = starts["3.4"]
        elif key == "3.4":
            end = starts["3.5"]
        elif key == "3.5":
            end = find("## 4. rekomendācija", after=start)
        elif key == "3.6":
            end = find("# A pielikums", after=start)
        else:
            end = next_index(start, lambda b: b.kind.startswith("h"))
        result[key] = (start, end)
    return result


def translate_source_notes(text: str) -> str:
    step_notes = {
        "1. Diagnose need (rank operators & identify binding constraint)": "1. Diagnosticēt vajadzību (sarindot pakalpojumu sniedzējus un identificēt saistošo ierobežojumu)",
        "2. Define feasible consolidation groupings (clusters)": "2. Noteikt iespējamas konsolidācijas grupas (kopas)",
        "3. Select consolidation model (shared services vs regional operator vs merger)": "3. Izvēlēties konsolidācijas modeli (kopīgi pakalpojumi, reģionāls pakalpojumu sniedzējs vai apvienošanās)",
        "4. Implementation & investment pathway (deliverability & sequencing)": "4. Īstenošanas un investīciju ceļš (īstenojamība un secība)",
    }
    stripped = text.strip()
    if stripped in step_notes:
        return step_notes[stripped]
    text = text.replace("Sources:", "Avoti:").replace("Source:", "Avots:")
    text = text.replace("Note:", "Piezīme:").replace("Notes:", "Piezīmes:")
    return text


def map_blocks_to_paragraphs(paragraphs: list[Paragraph], blocks: list[Block], start_index: int) -> None:
    """Map direct body paragraphs to the LV stream, retaining source elements."""
    cursor = start_index
    last_target: Paragraph | None = None
    for paragraph in paragraphs:
        if not paragraph.text.strip() and paragraph_has_drawing(paragraph):
            continue
        if not paragraph.text.strip():
            continue
        if paragraph.style.name == "Source & notes" and cursor < len(blocks):
            current_note = plain_text(block_text(blocks[cursor])).lower()
            if not re.match(r"^(?:avots|avoti|piezīme|piezīmes)\b|^(?:dz:|nt:|dw:|ww:)", current_note):
                set_paragraph_text(paragraph, translate_source_notes(paragraph.text))
                last_target = paragraph
                continue
        group, expected = source_kind(paragraph)
        short_para_caption = (
            group == "normal" and cursor < len(blocks)
            and blocks[cursor].kind == "caption" and len(paragraph.text.strip()) < 100
        )
        current_matches = cursor < len(blocks) and block_matches(blocks[cursor], expected)
        if current_matches:
            found = cursor
        elif group == "heading" or expected == "caption" or short_para_caption or (cursor < len(blocks) and blocks[cursor].kind not in {"caption", "h1", "h2", "h3", "h4", "h5", "h6"}):
            search_limit = 36 if group == "heading" else 1
            found = find_block(blocks, cursor, "caption" if short_para_caption else expected, search_limit)
        else:
            found = None
        if found is None:
            if paragraph.style.name == "Source & notes":
                set_paragraph_text(paragraph, translate_source_notes(paragraph.text))
                last_target = paragraph
            continue
        if found > cursor and last_target is not None:
            skipped = [plain_text(x.text) for x in blocks[cursor:found] if x.kind in {"para", "bullet", "ordered"}]
            if skipped:
                set_paragraph_text(last_target, f"{last_target.text} {' '.join(skipped)}")
        set_paragraph_text(paragraph, block_text(blocks[found]))
        last_target = paragraph
        cursor = found + 1


def map_box_paragraphs(cell: _Cell, lv_blocks: list[Block]) -> None:
    """Translate a one-cell box while leaving any drawings in place."""
    source_paragraphs = [p for p in cell.paragraphs if p.text.strip() or paragraph_has_drawing(p)]
    cursor = 0
    source_bullets = sum(1 for p in source_paragraphs if p.style.name.startswith("Bulleted"))
    lv_lists = sum(1 for b in lv_blocks if b.kind in {"bullet", "ordered"})
    first_list_extra = bool(lv_blocks and "3.4. ielikums" in plain_text(lv_blocks[0].text))
    seen_source_bullets = 0

    for paragraph in source_paragraphs:
        if not paragraph.text.strip() and paragraph_has_drawing(paragraph):
            continue
        group, expected = source_kind(paragraph)
        found = find_block(lv_blocks, cursor, expected, 8)
        if found is None:
            if paragraph.style.name == "Source & notes":
                set_paragraph_text(paragraph, translate_source_notes(paragraph.text))
            continue
        cursor = found
        text_parts = [block_text(lv_blocks[cursor])]
        cursor += 1
        if group == "list":
            seen_source_bullets += 1
            merge = 0
            if lv_lists > source_bullets:
                if first_list_extra and seen_source_bullets == 1:
                    merge = 1
                elif not first_list_extra and seen_source_bullets == source_bullets:
                    merge = 1
            if merge and cursor < len(lv_blocks) and lv_blocks[cursor].kind in {"bullet", "ordered"}:
                text_parts.append(block_text(lv_blocks[cursor]))
                cursor += 1
        set_paragraph_text(paragraph, " ".join(text_parts))


def map_table(table: Table, rows: list[list[str]]) -> None:
    if not rows:
        return
    for ri, row in enumerate(table.rows):
        if ri >= len(rows):
            break
        for ci, cell in enumerate(row.cells):
            if ci < len(rows[ri]):
                set_cell_text(cell, rows[ri][ci])


def table_blocks(blocks: list[Block]) -> list[Block]:
    return [b for b in blocks if b.kind == "table"]


def map_abbreviations(doc: Document, blocks: list[Block]) -> None:
    rows = blocks[1].rows or []
    acronym_rows = rows[1:] if rows and rows[0][0].lower() in {"saīsinājums", "abbr.", "abbreviation"} else rows
    direct = [Paragraph(el, doc) for el in doc.element.body.iterchildren() if el.tag == qn("w:p")]
    title_index = next(i for i, p in enumerate(direct) if p.text.strip() == "Abbreviations and acronyms")
    set_paragraph_text(direct[title_index], "Saīsinājumi un akronīmi")
    row_index = 0
    for paragraph in direct[title_index + 1 :]:
        if paragraph.style.name == "Heading 1" and "Water supply" in paragraph.text:
            break
        if not paragraph.text.strip() or row_index >= len(acronym_rows):
            continue
        abbr = paragraph.text.split("\t", 1)[0].strip()
        meaning = acronym_rows[row_index][1] if len(acronym_rows[row_index]) > 1 else ""
        set_paragraph_text(paragraph, f"{abbr}\t{plain_text(meaning)}")
        row_index += 1


def update_cover_and_toc(doc: Document) -> None:
    replacements = {
        "Organisation for Economic Co-operation and Development": "Ekonomiskās sadarbības un attīstības organizācija",
        "For Official Use": "Iekšējai lietošanai",
        "English - Or. English": "Latvian - Or. Latvian",
        "Fostering water resilience in Latvia": "Ūdens resursu noturības stiprināšana Latvijā",
        "Towards a long-term strategy for drinking water supply and wastewater management": "Ceļā uz ilgtermiņa dzeramā ūdensapgādes un notekūdeņu apsaimniekošanas stratēģiju",
        "Table of contents": "Satura rādītājs",
        "References": "Atsauces",
        "1 Water supply and sanitation service provision in Latvia": "1 Ūdensapgādes un kanalizācijas pakalpojumu sniegšana Latvijā",
        "Water supply and sanitation services: the state of play": "Ūdensapgādes un kanalizācijas pakalpojumi pašlaik",
        "Megatrends challenging service provision in Latvia": "Pakalpojumu sniegšanu Latvijā apgrūtinošās megatendences",
        "2 The governance of water supply and sanitation in Latvia": "2 Ūdensapgādes un kanalizācijas pārvaldība Latvijā",
        "Policymaking": "Politikas veidošana",
        "Service provision and infrastructure development": "Pakalpojumu sniegšana un infrastruktūras attīstība",
        "Economic regulation": "Ekonomiskais regulējums",
        "Environmental regulation": "Vides regulējums",
        "3 Towards a long-term strategy for water supply and sanitation": "3 Ceļā uz ilgtermiņa ūdensapgādes un kanalizācijas stratēģiju",
        "Recommendation 1: Strengthen national water governance through strategic vision, institutional capacity and policy coherence": "1. rekomendācija: Stiprināt valsts ūdens pārvaldību ar stratēģisku redzējumu, institucionālo kapacitāti un politikas saskaņotību",
        "Recommendation 1:": "1. rekomendācija:",
        "Recommendation 2: Align water services and infrastructure with territorial development": "2. rekomendācija: Saskaņot ūdensapgādes un kanalizācijas pakalpojumus un infrastruktūru ar teritoriālo attīstību",
        "Recommendation 3. Enhance the performance, financial sustainability and resilience of WSS service delivery": "3. rekomendācija. Uzlabot ŪK pakalpojumu sniegšanas darbības rezultātus, finansiālo ilgtspēju un noturību",
        "Recommendation 4. Secure sustainable funding and financing for water services": "4. rekomendācija. Nodrošināt ilgtspējīgu ūdensapgādes un kanalizācijas pakalpojumu finansējumu un finansēšanu",
        "Annex A. Assessment framework for WSS utility aggregation": "A pielikums. ŪK pakalpojumu sniedzēju apvienošanas novērtēšanas ietvars",
        "Step 1. Diagnose and classify operators: identifying where intervention is needed": "1. solis. Diagnosticēt un klasificēt pakalpojumu sniedzējus: noteikt, kur nepieciešama intervence",
        "Step 2. Feasible consolidation groupings: defining where consolidation can realistically occur": "2. solis. Iespējamas konsolidācijas grupas: noteikt, kur konsolidācija reāli var notikt",
        "Step 3. Select the appropriate consolidation model: matching solutions to problems": "3. solis. Izvēlēties atbilstošu konsolidācijas modeli: risinājumus saskaņot ar problēmām",
        "Step 2. Define feasible consolidation groupings (clusters): identify which operators could be grouped together": "2. solis. Noteikt iespējamas konsolidācijas grupas (kopas): identificēt, kurus pakalpojumu sniedzējus varētu apvienot",
        "Step 3. Select a consolidation model: shared services, regional operator or merger": "3. solis. Izvēlēties konsolidācijas modeli: kopīgi pakalpojumi, reģionāls pakalpojumu sniedzējs vai apvienošanās",
        "Step 4. Implementation and investment pathway: translating decisions into action": "4. solis. Īstenošanas un investīciju ceļš: lēmumu pārvēršana rīcībā",
        "This is a CONFIDENTIAL DRAFT and should not be shared.": "KONFIDENCIĀLS PROJEKTS, kas nav izplatāms.",
        "This document, as well as any data and map included herein, are without prejudice to the status of or sovereignty over any territory, to the delimitation of international frontiers and boundaries and to the name of any territory, city or area.": "Šis dokuments, kā arī tajā ietvertie dati un kartes neskar nevienas teritorijas statusu vai suverenitāti, starptautisko robežu un robežlīniju delimitāciju vai jebkuras teritorijas, pilsētas vai apgabala nosaukumu.",
    }
    abstract = (
        "Latvijas ūdensapgādes un kanalizācijas nozare atrodas pagrieziena punktā. Pēdējo divu desmitgažu laikā būtiskas investīcijas, ko lielā mērā atbalstīja Eiropas Savienības fondi, ir paplašinājušas piekļuvi dzeramā ūdensapgādes un notekūdeņu pakalpojumiem un uzlabojušas pakalpojumu kvalitāti. Šobrīd nozare saskaras ar jauniem izaicinājumiem: novecojošu infrastruktūru, stingrākām regulatīvajām prasībām, demogrāfisko lejupslīdi, pieaugošiem klimata un drošības riskiem un arvien ierobežotāku publisko un ES finansējumu. Šīs tendences apdraud pakalpojumu sniegšanas finansiālo ilgtspēju un ilgtermiņa dzīvotspēju. Pamatojoties uz politikas dialogu ar Klimata un enerģētikas ministriju un vairāk nekā 100 ieinteresētajām pusēm, ziņojumā analizēta ūdensapgādes un kanalizācijas pakalpojumu pārvaldība un finansēšana Latvijā. Tajā noteiktas galvenās pārvaldības nepilnības, kas kavē pāreju uz noturīgu, finansiāli ilgtspējīgu un konkurētspējīgu nozari, un sniegtas rekomendācijas daudzlīmeņu pārvaldības stiprināšanai, pakalpojumu sniegšanas un regulējuma uzlabošanai, kā arī finansējuma un finansēšanas mobilizēšanai ilgtermiņa ūdensnoturībai."
    )
    # TOC entries are split across hyperlink runs and their page-number
    # fields. Replace the text runs as a unit while leaving PAGEREF fields
    # intact, so page numbers remain in the original layout.
    for element in doc.element.body.iter():
        if element.tag != qn("w:p"):
            continue
        paragraph = Paragraph(element, doc)
        nodes = text_nodes(element)
        field_nodes = field_result_nodes(paragraph)
        visible_nodes = [node for node in nodes if node not in field_nodes]
        joined = "".join(node.text or "" for node in visible_nodes)
        value = joined
        for old, new in replacements.items():
            value = value.replace(old, new)
        if value == joined or not visible_nodes:
            continue
        target = next((node for node in visible_nodes if node.text and node.text.strip()), visible_nodes[0])
        target.text = value
        target.set(XML_SPACE, "preserve")
        for node in visible_nodes:
            if node is target:
                continue
            node.text = ""
            node.set(XML_SPACE, "preserve")
    for element in doc.element.body.iter():
        if element.tag != qn("w:t") or not element.text:
            continue
        value = element.text
        for old, new in replacements.items():
            value = value.replace(old, new)
        element.text = value
    # The confidential-draft sentence is split across separately formatted
    # runs, so it cannot be replaced by the run-level substitutions above.
    for element in doc.element.body.iter():
        if element.tag != qn("w:p"):
            continue
        paragraph = Paragraph(element, doc)
        if "This is a CONFIDENTIAL DRAFT and should not be shared." not in paragraph.text:
            continue
        nodes = text_nodes(element)
        nonempty = [node for node in nodes if node.text and node.text.strip()]
        if len(nonempty) >= 2:
            nonempty[0].text = ""
            nonempty[1].text = "KONFIDENCIĀLS PROJEKTS"
            for node in nonempty[2:-1]:
                node.text = ""
            nonempty[-1].text = ", kas nav izplatāms."
        else:
            set_paragraph_text(paragraph, "KONFIDENCIĀLS PROJEKTS, kas nav izplatāms.")
    for element in doc.element.body.iter():
        if element.tag != qn("w:p"):
            continue
        paragraph = Paragraph(element, doc)
        if paragraph.style.name == "CoverInformation" and "Latvia’s water supply" in paragraph.text:
            set_paragraph_text(paragraph, abstract)


def update_footnotes(path: Path, translation_blocks: list[Block]) -> None:
    definitions: dict[str, str] = {}
    for block in translation_blocks:
        match = re.match(r"\[\^([0-9]+)\]:\s*(.*)$", block.text)
        if match:
            definitions[match.group(1)] = plain_text(match.group(2))
    if not definitions:
        return
    temp = path.with_suffix(".tmp.docx")
    # python-docx does not expose footnotes and may rewrite their part while
    # saving. Use the source part as the structural base, then replace only
    # the visible footnote text with the Latvian definitions.
    with zipfile.ZipFile(SOURCE, "r") as zsrc, zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zsrc.read("word/footnotes.xml") if item.filename == "word/footnotes.xml" else zin.read(item.filename)
            if item.filename == "word/footnotes.xml":
                root = etree.fromstring(data)
                for footnote in root:
                    if footnote.tag != f"{{{W}}}footnote":
                        continue
                    ident = footnote.get(f"{{{W}}}id")
                    if ident not in definitions:
                        continue
                    nodes = [node for node in footnote.iter() if node.tag == f"{{{W}}}t"]
                    if not nodes:
                        paragraph = next((node for node in footnote if node.tag == f"{{{W}}}p"), None)
                        if paragraph is None:
                            continue
                        run = etree.SubElement(paragraph, f"{{{W}}}r")
                        nodes = [etree.SubElement(run, f"{{{W}}}t")]
                    nodes[0].text = " " + definitions[ident]
                    nodes[0].set(XML_SPACE, "preserve")
                    for node in nodes[1:]:
                        node.text = ""
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            zout.writestr(item, data)
    temp.replace(path)


def main() -> None:
    if not SOURCE.is_file():
        raise FileNotFoundError(SOURCE)
    for name in ORDER:
        if not (TRANSLATION_DIR / name).is_file():
            raise FileNotFoundError(TRANSLATION_DIR / name)

    blocks = parse_markdown()
    box_ranges = identify_box_ranges(blocks)
    excluded: set[int] = set()
    for start, end in box_ranges.values():
        excluded.update(range(start, end))
    excluded.update(i for i, block in enumerate(blocks) if block.kind in {"table", "footnote"})
    excluded.add(0)  # the abbreviations heading is handled separately
    # The Markdown keeps this label as a standalone heading, while the source
    # document contains it inside Box 2.4. The box caption and body already
    # carry the text, so it must not enter the direct body stream twice.
    excluded.update(
        i for i, block in enumerate(blocks)
        if block.kind == "h6" and plain_text(block.text).startswith("Tarifu metodika ūdenssaimniecības")
    )
    main_blocks = [block for i, block in enumerate(blocks) if i not in excluded]
    # The source has one paragraph for the climate-impact discussion, while
    # the Markdown translation splits its final sentence into a separate
    # paragraph before Figure 1.9. Merge those two translation blocks so the
    # following figure caption remains aligned with the source template.
    for i in range(len(main_blocks) - 1):
        if (
            main_blocks[i].kind == "para"
            and main_blocks[i + 1].kind == "para"
            and plain_text(main_blocks[i].text).startswith("Ceturtais faktors")
            and plain_text(main_blocks[i + 1].text).startswith("Klimata riski")
        ):
            main_blocks[i].text = f"{main_blocks[i].text} {main_blocks[i + 1].text}"
            del main_blocks[i + 1]
            break
    tables = table_blocks(blocks)
    abbreviation_table = tables[0]
    report_tables = tables[1:]

    doc = Document(SOURCE)
    update_cover_and_toc(doc)
    map_abbreviations(doc, [blocks[0], abbreviation_table])

    body_paragraphs = [Paragraph(el, doc) for el in doc.element.body.iterchildren() if el.tag == qn("w:p")]
    start = next(i for i, paragraph in enumerate(body_paragraphs) if paragraph.style.name == "Heading 1" and "Water supply" in paragraph.text)
    map_blocks_to_paragraphs(body_paragraphs[start:], main_blocks, 0)

    source_table_index = 0
    box_keys = {
        1: "2.1", 2: "2.2", 4: "2.3", 5: "2.4", 6: "2.5",
        7: "3.1", 10: "3.2", 11: "3.3", 12: "3.4", 13: "3.5", 14: "3.6",
    }
    for table_number, table in enumerate(doc.tables):
        if len(table.rows) == 1 and len(table.columns) == 1:
            key = box_keys.get(table_number)
            if key is None:
                continue
            start_i, end_i = box_ranges[key]
            box_blocks = [b for b in blocks[start_i:end_i] if b.kind != "table"]
            map_box_paragraphs(table.cell(0, 0), box_blocks)
            for nested in table.cell(0, 0).tables:
                if source_table_index >= len(report_tables):
                    raise RuntimeError("Not enough translated table blocks")
                map_table(nested, report_tables[source_table_index].rows or [])
                source_table_index += 1
        else:
            if source_table_index >= len(report_tables):
                raise RuntimeError("Not enough translated table blocks")
            map_table(table, report_tables[source_table_index].rows or [])
            source_table_index += 1
    if source_table_index != len(report_tables):
        raise RuntimeError(f"Translated tables used {source_table_index}/{len(report_tables)}")

    doc.core_properties.title = "Ūdens resursu noturības stiprināšana Latvijā"
    doc.core_properties.subject = "OECD ziņojuma tulkojums latviešu valodā"
    doc.core_properties.language = "lv-LV"
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    update_footnotes(OUTPUT, blocks)
    print(f"{OUTPUT}\nmedia preserved from source template; tables translated: {source_table_index}")


if __name__ == "__main__":
    main()
