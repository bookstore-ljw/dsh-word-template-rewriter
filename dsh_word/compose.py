"""Compose a document from source content using the template as a style set.

The template contributes formatting only: page setup, headers/footers, styles,
caption/equation conventions.  The source contributes every heading, paragraph,
equation, figure and table, in order.  Equations, figures and tables are numbered
sequentially in the template's language.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Emu
from docx.text.paragraph import Paragraph

from .autoplan import _has_image, _save_images
from .media import image_extent
from .omml import extract_math, math_nodes

_SECTPR = qn("w:sectPr")


@dataclass
class ComposeResult:
    out_path: str
    headings: int = 0
    paragraphs: int = 0
    equations: int = 0
    figures: int = 0
    tables: int = 0
    mapping: list[dict] = field(default_factory=list)


def _style_names(template) -> dict:
    names = {style.name for style in template.styles}
    def pick(*candidates):
        for name in candidates:
            if name in names:
                return name
        return None
    heading = {}
    for level in range(1, 5):
        heading[level] = pick(f"Heading {level}", f"标题 {level}", f"标题{level}")
    body = pick("Normal", "正文", "Body Text")
    caption = next((n for n in names if "caption" in n.lower() or "题注" in n), None)
    return {
        "title": pick("Title", "标题"),
        "heading": heading,
        "body": body or "Normal",
        "caption": caption or body or "Normal",
    }


def _clear_body(document) -> None:
    for child in list(document.element.body.iterchildren()):
        if child.tag == _SECTPR:
            continue
        child.getparent().remove(child)


def _usable_width(document) -> int:
    section = document.sections[0]
    return int(section.page_width - section.left_margin - section.right_margin)


def _fit_width(document, width: int) -> int:
    available = _usable_width(document)
    if available and width > available:
        return available
    return width


def _source_blocks(document, assets: Path) -> list[dict]:
    blocks: list[dict] = []
    paragraph_index = -1
    table_index = -1
    image_index = 0
    for child in document.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            paragraph_index += 1
            paragraph = Paragraph(child, document)
            if math_nodes(paragraph):
                blocks.append({"kind": "equation", "paragraph": paragraph_index})
            elif _has_image(paragraph):
                for saved in _save_images(paragraph, document, assets, image_index + 1):
                    image_index += 1
                    blocks.append({"kind": "figure", "path": str(saved)})
            elif paragraph.text.strip():
                style = paragraph.style.name if paragraph.style is not None else "Normal"
                level = 0
                lowered = style.lower()
                if lowered.startswith("heading") or style.startswith("标题"):
                    digits = [c for c in style if c.isdigit()]
                    level = int(digits[0]) if digits else 1
                blocks.append(
                    {
                        "kind": "heading" if level else "text",
                        "level": level,
                        "text": paragraph.text.strip(),
                        "style": style,
                    }
                )
        elif tag == "tbl":
            table_index += 1
            blocks.append({"kind": "table", "table": table_index})
    return blocks


def _equation_style(document, fallback: str) -> str:
    """Use the template's own equation paragraph style for equations and their notes."""
    for paragraph in document.paragraphs:
        if math_nodes(paragraph) and paragraph.style is not None:
            return paragraph.style.name
    return fallback


def _caption_language(template) -> bool:
    text = "\n".join(p.text for p in template.paragraphs)
    return any("\u4e00" <= ch <= "\u9fff" for ch in text)


_CAPTION_RE = __import__("re").compile(
    r"^(图|表|figure|table|fig\.?)\s*[0-9一二三四五六七八九十]+([\s\.\-—:：]|$)",
    __import__("re").IGNORECASE,
)


def _looks_like_caption(text: str, kind: str) -> bool:
    return bool(_CAPTION_RE.match(text.strip()))


def _add_caption(document, style: str, label: str) -> Paragraph:
    paragraph = document.add_paragraph(style=style)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run(label)
    return paragraph


def _equation_paragraph(document, style: str, element, number: int, width: int) -> Paragraph:
    paragraph = document.add_paragraph(style=style)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        paragraph.paragraph_format.tab_stops.add_tab_stop(Emu(width), WD_TAB_ALIGNMENT.RIGHT)
    except Exception:  # noqa: BLE001 - tab stops are best effort
        pass
    if element is not None:
        paragraph._p.append(copy.deepcopy(element))
    paragraph.add_run(f"\t({number})")
    return paragraph


def compose(
    template_path: str | Path,
    source_path: str | Path,
    out_path: str | Path,
    steps_dir: str | Path | None = None,
    on_op=None,
) -> ComposeResult:
    template_path = Path(template_path)
    source_path = Path(source_path)
    out_path = Path(out_path)
    document = Document(str(template_path))
    source = Document(str(source_path))
    assets = out_path.parent / "assets"
    blocks = _source_blocks(source, assets)
    styles = _style_names(document)
    equation_style = _equation_style(document, styles["body"])
    width = _usable_width(document)
    chinese = _caption_language(document)

    _clear_body(document)
    result = ComposeResult(out_path=str(out_path))
    steps_root = Path(steps_dir) if steps_dir else None
    if steps_root is not None:
        steps_root.mkdir(parents=True, exist_ok=True)

    figure_no = 0
    table_no = 0
    equation_no = 0

    def emit(index: int, kind: str, detail: str) -> None:
        if steps_root is not None:
            step = steps_root / f"{out_path.stem}.step.{index:02d}.docx"
            document.save(str(step))
            from .fill import _save_highlight_copy

            _save_highlight_copy(str(step), steps_root, out_path.stem, index, {"kind": "body", "paragraph": len(document.paragraphs) - 1})
        if on_op is not None:
            on_op(
                {
                    "op": index,
                    "kind": kind,
                    "slot": f"{kind}.c{index}",
                    "location": {"kind": "body", "paragraph": len(document.paragraphs) - 1},
                    "detail": detail,
                    "step": str(steps_root / f"{out_path.stem}.step.{index:02d}.docx") if steps_root else "",
                }
            )

    index = 0
    skipped: set[int] = set()
    for position, block in enumerate(blocks):
        if position in skipped:
            continue
        index += 1
        if block["kind"] == "heading":
            level = min(max(block.get("level", 1), 1), 4)
            style = styles["heading"].get(level) or styles["body"]
            paragraph = document.add_paragraph(style=style)
            paragraph.add_run(block["text"])
            result.headings += 1
            emit(index, "heading", block["text"][:60])
        elif block["kind"] == "text":
            paragraph = document.add_paragraph(style=styles["body"])
            paragraph.add_run(block["text"])
            result.paragraphs += 1
            emit(index, "text", block["text"][:60])
        elif block["kind"] == "equation":
            equation_no += 1
            element = extract_math(source.paragraphs[block["paragraph"]])
            _equation_paragraph(document, equation_style, element, equation_no, width)
            result.equations += 1
            emit(index, "equation", f"公式 ({equation_no})")
        elif block["kind"] == "figure":
            figure_no += 1
            paragraph = document.add_paragraph(style=styles["body"])
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            width_emu, _height = image_extent(block["path"])
            paragraph.add_run().add_picture(block["path"], width=Emu(_fit_width(document, width_emu)))
            caption = None
            if position + 1 < len(blocks) and blocks[position + 1]["kind"] == "text":
                candidate = blocks[position + 1]["text"]
                if _looks_like_caption(candidate, "figure"):
                    caption = candidate
                    skipped.add(position + 1)
            label = caption or (f"图 {figure_no}" if chinese else f"Figure {figure_no}")
            _add_caption(document, styles["caption"], label)
            result.figures += 1
            emit(index, "figure", label[:60])
        else:
            table_no += 1
            table = source.tables[block["table"]]
            clone = copy.deepcopy(table._tbl)
            caption = None
            if position + 1 < len(blocks) and blocks[position + 1]["kind"] == "text":
                candidate = blocks[position + 1]["text"]
                if _looks_like_caption(candidate, "table"):
                    caption = candidate
                    skipped.add(position + 1)
            label = caption or (f"表 {table_no}" if chinese else f"Table {table_no}")
            _add_caption(document, styles["caption"], label)
            document.element.body.append(clone)
            result.tables += 1
            emit(index, "table", label[:60])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(out_path))
    return result
