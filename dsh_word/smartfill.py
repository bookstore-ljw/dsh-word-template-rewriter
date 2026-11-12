"""In-place fill for templates without ``{{ }}`` markers.

The template keeps its structure: headings, paragraph order and every paragraph's
``pPr`` stay where they are.  Source blocks are written into the template's body
paragraphs one by one, so the surrounding formatting (style, indentation, spacing)
survives.  Extra source blocks are appended with the last body paragraph as the
format donor; unused template paragraphs are emptied rather than removed.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

MATH_FONTS = {'Symbol', 'Cambria Math', 'CambriaMath', 'MS Math'}


def _strip_math_font(run) -> None:
    rpr = run._element.rPr
    if rpr is None:
        return
    fonts = rpr.find(qn('w:rFonts'))
    if fonts is None:
        return
    names = {fonts.get(qn(f'w:{key}')) for key in ('ascii', 'hAnsi', 'eastAsia', 'cs')}
    if names & MATH_FONTS:
        rpr.remove(fonts)
from docx.shared import Emu
from docx.text.paragraph import Paragraph

from .autoplan import _has_image, _save_images
from .fill import _save_highlight_copy
from .media import image_extent
from .omml import extract_math, math_nodes


@dataclass
class SmartFillResult:
    out_path: str
    markers: int = 0
    appended: int = 0
    tables: int = 0
    mapping: list[dict] = field(default_factory=list)


def has_markers(path: str | Path) -> bool:
    from .slots import find_markers

    document = Document(str(path))
    return any(find_markers(paragraph) for paragraph in document.paragraphs)


def _is_heading(paragraph: Paragraph) -> bool:
    style = paragraph.style.name if paragraph.style is not None else ""
    return bool(style) and style.lower().startswith(("heading", "标题", "title"))


def _clear_runs(paragraph: Paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag == qn("w:pPr"):
            continue
        paragraph._p.remove(child)


def _set_text(paragraph: Paragraph, text: str) -> None:
    """Replace the paragraph text but keep its paragraph/character formatting."""
    protos = [copy.deepcopy(run._r) for run in paragraph.runs[:1]]
    _clear_runs(paragraph)
    if protos:
        run_element = protos[0]
        for child in list(run_element):
            if child.tag == qn("w:t"):
                run_element.remove(child)
        paragraph._p.append(run_element)
        run = paragraph.runs[-1]
    else:
        run = paragraph.add_run()
    run.text = text


def _fit_width(document, width: int) -> int:
    try:
        section = document.sections[0]
        available = section.page_width - section.left_margin - section.right_margin
        if available and width > available:
            return int(available)
    except Exception:  # noqa: BLE001
        pass
    return width


def _set_image(paragraph: Paragraph, path: str, width: int) -> None:
    _clear_runs(paragraph)
    paragraph.add_run().add_picture(path, width=Emu(width))


def _set_formula(paragraph: Paragraph, element) -> None:
    _clear_runs(paragraph)
    if element is not None:
        paragraph._p.append(copy.deepcopy(element))


def _append_like(document, reference: Paragraph) -> Paragraph:
    paragraph = document.add_paragraph()
    if reference._p.pPr is not None:
        paragraph._p.insert(0, copy.deepcopy(reference._p.pPr))
    return paragraph


def _source_blocks(document, asset_dir: Path) -> list[dict]:
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
                blocks.append({"kind": "formula", "paragraph": paragraph_index})
            elif _has_image(paragraph):
                for saved in _save_images(paragraph, document, asset_dir, image_index + 1):
                    image_index += 1
                    blocks.append({"kind": "image", "path": str(saved)})
            elif paragraph.text.strip():
                blocks.append({"kind": "text", "text": paragraph.text.strip()})
        elif tag == "tbl":
            table_index += 1
            blocks.append({"kind": "table", "table": table_index})
    return blocks


def smart_fill(
    template_path: str | Path,
    source_path: str | Path,
    out_path: str | Path,
    asset_dir: str | Path | None = None,
    steps_dir: str | Path | None = None,
    on_op=None,
) -> SmartFillResult:
    template_path = Path(template_path)
    source_path = Path(source_path)
    out_path = Path(out_path)
    assets = Path(asset_dir) if asset_dir else out_path.parent / "assets"

    template = Document(str(template_path))
    source = Document(str(source_path))
    blocks = _source_blocks(source, assets)
    result = SmartFillResult(out_path=str(out_path))

    targets = [
        index
        for index, paragraph in enumerate(template.paragraphs)
        if not _is_heading(paragraph)
    ]
    reference = template.paragraphs[targets[-1]] if targets else template.paragraphs[-1]
    queue = list(targets)
    steps_root = Path(steps_dir) if steps_dir else None
    if steps_root is not None:
        steps_root.mkdir(parents=True, exist_ok=True)

    def step(index: int, block: dict, detail: str) -> None:
        if steps_root is not None:
            path = steps_root / f"{out_path.stem}.step.{index:02d}.docx"
            template.save(str(path))
            _save_highlight_copy(
                str(path),
                steps_root,
                out_path.stem,
                index,
                {"kind": "body", "paragraph": len(template.paragraphs) - 1},
            )
        if on_op is not None:
            on_op(
                {
                    "op": index,
                    "kind": block["kind"],
                    "slot": f"{block['kind']}.s{index}",
                    "location": {"kind": "body", "paragraph": len(template.paragraphs) - 1},
                    "detail": detail,
                    "step": str(steps_root / f"{out_path.stem}.step.{index:02d}.docx")
                    if steps_root
                    else "",
                }
            )

    for index, block in enumerate(blocks, start=1):
        if queue:
            paragraph = template.paragraphs[queue.pop(0)]
        else:
            paragraph = _append_like(template, reference)
            result.appended += 1

        if block["kind"] == "text":
            _set_text(paragraph, block["text"])
            detail = block["text"][:60]
        elif block["kind"] == "image":
            width, _height = image_extent(block["path"])
            _set_image(paragraph, block["path"], _fit_width(template, width))
            detail = Path(block["path"]).name
        elif block["kind"] == "formula":
            _set_formula(paragraph, extract_math(source.paragraphs[block["paragraph"]]))
            detail = f"公式 {index}"
        else:
            table_element = copy.deepcopy(source.tables[block["table"]]._tbl)
            paragraph._p.addnext(table_element)
            result.tables += 1
            detail = f"表格 {block['table'] + 1}"

        result.markers += 1
        result.mapping.append(
            {
                "kind": block["kind"],
                "slot": f"{block['kind']}.s{index}",
                "from": detail,
                "paragraph": index - 1,
            }
        )
        step(index, block, detail)

    # Unused sample paragraphs are emptied, not removed: their formatting stays.
    for index in queue:
        if index < len(template.paragraphs):
            _clear_runs(template.paragraphs[index])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    template.save(str(out_path))
    return result
