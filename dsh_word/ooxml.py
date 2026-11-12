"""Read-only OOXML helpers: document order walking, formatting signatures.

The fill worker mutates a document; everything in this module only observes it.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

W_P = qn("w:p")
W_TBL = qn("w:tbl")
W_TR = qn("w:tr")
W_TC = qn("w:tc")
W_R = qn("w:r")
W_RPR = qn("w:rPr")
W_PPR = qn("w:pPr")
W_DRAWING = qn("w:drawing")
W_PICT = qn("w:pict")
M_OMATH = qn("m:oMath")
M_OMATHPARA = qn("m:oMathPara")


@dataclass(frozen=True)
class Location:
    """Where a paragraph lives, stable enough to quote in a report."""

    kind: str  # "body" | "table"
    index: int
    table: int | None = None
    row: int | None = None
    cell: int | None = None

    def as_dict(self) -> dict:
        data: dict = {"kind": self.kind, "paragraph": self.index}
        if self.table is not None:
            data.update(table=self.table, row=self.row, cell=self.cell)
        return data

    def label(self) -> str:
        if self.table is None:
            return f"body.p[{self.index}]"
        return f"table[{self.table}].row[{self.row}].cell[{self.cell}].p[{self.index}]"


def iter_paragraphs(doc) -> Iterator[tuple[Location, Paragraph]]:
    """Yield every paragraph in document order, body and tables included."""
    counters = {"table": 0}
    yield from _walk(doc.element.body, doc, "body", counters, None, None, None)


def _walk(container, doc, kind, counters, table, row, cell):
    index = 0
    for child in container.iterchildren():
        if child.tag == W_P:
            yield Location(kind, index, table, row, cell), Paragraph(child, doc)
            index += 1
        elif child.tag == W_TBL:
            table_index = counters["table"]
            counters["table"] += 1
            for row_index, tr in enumerate(child.findall(W_TR)):
                for cell_index, tc in enumerate(tr.findall(W_TC)):
                    yield from _walk(
                        tc, doc, "table", counters, table_index, row_index, cell_index
                    )


def body_paragraphs(doc) -> list[Paragraph]:
    return [Paragraph(p, doc) for p in doc.element.body.iterchildren() if p.tag == W_P]


def paragraph_text(paragraph: Paragraph) -> str:
    return "".join(run.text for run in paragraph.runs)


def document_text(doc) -> str:
    return "\n".join(paragraph_text(p) for _, p in iter_paragraphs(doc))


def is_math_element(element) -> bool:
    return element.tag in (M_OMATH, M_OMATHPARA)


def is_drawing_run(run) -> bool:
    element = run._r
    return element.find(".//" + W_DRAWING) is not None or element.find(
        ".//" + W_PICT
    ) is not None


def run_signature(run) -> dict:
    """The character formatting that must survive a copy, as plain data."""
    rPr = run._r.find(W_RPR)
    if rPr is None:
        return {}
    sig: dict = {}
    fonts = rPr.find(qn("w:rFonts"))
    if fonts is not None:
        for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            value = fonts.get(qn(attribute))
            if value is not None:
                sig[attribute.split(":")[1]] = value
    for tag, key in (("w:sz", "sz"), ("w:szCs", "szCs"), ("w:color", "color")):
        element = rPr.find(qn(tag))
        if element is not None:
            sig[key] = element.get(qn("w:val"))
    for tag, key in (("w:b", "bold"), ("w:i", "italic"), ("w:u", "underline")):
        element = rPr.find(qn(tag))
        if element is not None:
            value = element.get(qn("w:val"))
            sig[key] = value not in ("0", "false", "off")
    vert = rPr.find(qn("w:vertAlign"))
    if vert is not None:
        sig["vertAlign"] = vert.get(qn("w:val"))
    return sig


def paragraph_signature(paragraph: Paragraph) -> dict:
    """Paragraph level formatting: style, alignment, spacing, indentation."""
    pPr = paragraph._p.find(W_PPR)
    if pPr is None:
        return {}
    sig: dict = {}
    style = pPr.find(qn("w:pStyle"))
    if style is not None:
        sig["style"] = style.get(qn("w:val"))
    for tag, key in (
        ("w:jc", "align"),
        ("w:outlineLvl", "outline"),
        ("w:keepNext", "keepNext"),
    ):
        element = pPr.find(qn(tag))
        if element is not None:
            sig[key] = element.get(qn("w:val")) or True
    spacing = pPr.find(qn("w:spacing"))
    if spacing is not None:
        for key in ("before", "after", "line", "beforeLines", "afterLines", "lineRule"):
            value = spacing.get(qn("w:" + key))
            if value is not None:
                sig[key] = value
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        for key in ("left", "right", "firstLine", "firstLineChars", "hanging"):
            value = ind.get(qn("w:" + key))
            if value is not None:
                sig[key] = value
    num = pPr.find(qn("w:numPr"))
    if num is not None:
        ilvl = num.find(qn("w:ilvl"))
        num_id = num.find(qn("w:numId"))
        sig["numPr"] = (
            ilvl.get(qn("w:val")) if ilvl is not None else None,
            num_id.get(qn("w:val")) if num_id is not None else None,
        )
    return sig
