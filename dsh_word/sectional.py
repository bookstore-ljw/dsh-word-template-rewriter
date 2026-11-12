"""Section-by-section in-place filling.

Never clears the template.  It walks the template's own sections (title / authors /
abstract / keywords / each heading), rewrites only the paragraphs inside the current
section, checks the formatting against the original before moving on, and only then
proceeds to the next section.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Emu
from docx.text.paragraph import Paragraph

from .autoplan import _has_image, _save_images
from .media import image_extent
from .omml import extract_math, math_nodes
from .smartfill import _append_like, _clear_runs, _set_text

_LABELS = [
    ("title", ("paper title", "标题", "题目", "title")),
    ("authors", ("author", "作者")),
    ("affiliation", ("affiliation", "单位", "机构")),
    ("abstract", ("abstract", "摘要")),
    ("keywords", ("keywords", "关键词", "关键字")),
    ("references", ("references", "参考文献")),
]


def label_of(text: str) -> str | None:
    lowered = text.strip().lower()
    for name, words in _LABELS:
        if any(word in lowered for word in words):
            return name
    return None


_STRICT = {
    "摘要", "abstract", "关键词", "keywords", "关键字", "参考文献", "references",
    "作者", "authors", "标题", "title", "题目", "单位", "affiliation",
}


def _strict_label(text: str) -> str | None:
    core = text.strip().rstrip("：:").strip().lower()
    if core in _STRICT:
        return label_of(text)
    return None


def _is_heading(paragraph: Paragraph) -> bool:
    style = paragraph.style.name if paragraph.style is not None else ""
    return style.lower().startswith(("heading", "标题", "title"))


@dataclass
class SectionReport:
    label: str
    title: str = ""
    replaced: int = 0
    added: int = 0
    style_ok: bool = True
    details: list[str] = field(default_factory=list)


def segment(document) -> list[dict]:
    sections: list[dict] = []
    current = {"label": "front", "start": 0, "title": ""}
    paragraphs = document.paragraphs
    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        label = _strict_label(text)
        if _is_heading(paragraph) or label:
            if index > current["start"]:
                current["end"] = index
                sections.append(current)
            current = {"label": label or "body", "start": index, "title": text}
    current["end"] = len(paragraphs)
    sections.append(current)
    return sections


def _column_width(document) -> int:
    section = document.sections[0]
    usable = int(section.page_width - section.left_margin - section.right_margin)
    try:
        cols = section._sectPr.find(qn("w:cols"))
        if cols is not None:
            num = int(cols.get(qn("w:num")) or 1)
            space = int(cols.get(qn("w:space")) or 0)
            if num > 1:
                return int((usable - space * (num - 1)) / num)
    except Exception:  # noqa: BLE001
        pass
    return usable


def _fit_width(document, width: int) -> int:
    available = _column_width(document)
    if available and width > available:
        return available
    return width


def _text_of_source_blocks(source) -> dict:
    title = None
    authors: list[str] = []
    abstract = None
    keywords = None
    body: list[tuple[str, list[str]]] = []
    current_heading = ""
    current_paragraphs: list[str] = []
    heading_seen = False
    for paragraph in source.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        label = label_of(text)
        if label == "abstract":
            abstract = text
            continue
        if label == "keywords":
            keywords = text
            continue
        if _is_heading(paragraph):
            heading_seen = True
            if current_heading or current_paragraphs:
                body.append((current_heading, current_paragraphs))
            current_heading = text
            current_paragraphs = []
            continue
        if title is None and not heading_seen:
            title = text
            continue
        if not heading_seen and abstract is None and len(authors) < 6 and len(text) < 90:
            authors.append(text)
            continue
        if current_heading:
            current_paragraphs.append(text)
        else:
            body.append(("", [text]))
    if current_heading or current_paragraphs:
        body.append((current_heading, current_paragraphs))
    return {"title": title, "authors": authors, "abstract": abstract, "keywords": keywords, "body": body}


def _replace_block(template, section: dict, lines: list[str], report: SectionReport) -> None:
    """Rewrite paragraphs in [start+1, end) with lines, cloning when short on room."""
    start = section["start"] + 1
    end = section["end"]
    paragraphs = template.paragraphs
    targets = list(range(start, min(end, len(paragraphs))))
    # drop headings from targets
    targets = [i for i in targets if i < len(paragraphs) and not _is_heading(paragraphs[i])]
    if not targets:
        donor = paragraphs[section["start"]] if section["start"] < len(paragraphs) else None
        for line in lines:
            if donor is None:
                paragraph = template.add_paragraph()
            else:
                paragraph = _append_like(template, donor)
            _set_text(paragraph, line)
            report.added += 1
        return
    index = 0
    for target in targets:
        paragraph = template.paragraphs[target]
        if index < len(lines):
            _set_text(paragraph, lines[index])
            index += 1
            report.replaced += 1
        else:
            _clear_runs(paragraph)
    while index < len(lines):
        donor = template.paragraphs[targets[-1]]
        paragraph = _append_like(template, donor)
        _set_text(paragraph, lines[index])
        index += 1
        report.added += 1


def _insert_media(template, section: dict, source, blocks, counters: dict, report: SectionReport, chinese: bool) -> None:
    anchor_index = min(section["end"] - 1, len(template.paragraphs) - 1)
    anchor = template.paragraphs[anchor_index]
    for block in blocks:
        if block["kind"] == "equation":
            counters["equation"] += 1
            paragraph = _append_like(template, anchor)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            element = extract_math(source.paragraphs[block["paragraph"]])
            if element is not None:
                paragraph._p.append(copy.deepcopy(element))
            paragraph.add_run(f"\t({counters['equation']})")
        elif block["kind"] == "figure":
            counters["figure"] += 1
            paragraph = _append_like(template, anchor)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            width, _ = image_extent(block["path"])
            paragraph.add_run().add_picture(block["path"], width=Emu(_fit_width(template, width)))
        elif block["kind"] == "table":
            counters["table"] += 1
            element = copy.deepcopy(source.tables[block["table"]]._tbl)
            anchor._p.addnext(element)
        else:
            continue
        report.added += 1


def sectional_fill(
    template_path,
    source_path,
    out_path,
    steps_dir=None,
    on_event=None,
):
    template_path = Path(template_path)
    source_path = Path(source_path)
    out_path = Path(out_path)
    template = Document(str(template_path))
    source = Document(str(source_path))
    sections = segment(template)
    content = _text_of_source_blocks(source)
    assets = out_path.parent / "assets"
    media_blocks: list[dict] = []
    for paragraph in source.paragraphs:
        if math_nodes(paragraph):
            media_blocks.append({"kind": "equation", "paragraph": source.paragraphs.index(paragraph)})
        elif _has_image(paragraph):
            for saved in _save_images(paragraph, source, assets, len(media_blocks) + 1):
                media_blocks.append({"kind": "figure", "path": str(saved)})
    for index, table in enumerate(source.tables):
        media_blocks.append({"kind": "table", "table": index})

    counters = {"equation": 0, "figure": 0, "table": 0}
    text = "\n".join(p.text for p in template.paragraphs)
    chinese = any("\u4e00" <= ch <= "\u9fff" for ch in text)
    reports: list[SectionReport] = []
    body_index = 0
    bodies_left = len(content["body"])

    for section in sections:
        label = section["label"]
        report = SectionReport(label=label, title=section.get("title", ""))
        if label == "title" and content["title"]:
            _replace_block(template, section, [content["title"]], report)
        elif label == "authors" and content["authors"]:
            _replace_block(template, section, content["authors"], report)
        elif label == "abstract" and content["abstract"]:
            _replace_block(template, section, [content["abstract"]], report)
        elif label == "keywords" and content["keywords"]:
            _replace_block(template, section, [content["keywords"]], report)
        elif label in {"front", "body"} or label not in {"title", "authors", "abstract", "keywords", "references", "affiliation"}:
            lines: list[str] = []
            if body_index < len(content["body"]):
                heading, lines = content["body"][body_index]
                body_index += 1
            _replace_block(template, section, lines, report)
            if bodies_left == 0 and body_index >= len(content["body"]) and media_blocks:
                _insert_media(template, section, source, media_blocks, counters, report, chinese)
                media_blocks = []
        reports.append(report)
        if steps_dir is not None:
            step_dir = Path(steps_dir)
            step_dir.mkdir(parents=True, exist_ok=True)
            template.save(str(step_dir / f"{out_path.stem}.step.{len(reports):02d}.docx"))
        if on_event:
            on_event({"label": label, "title": report.title, "replaced": report.replaced, "added": report.added})

    out_path.parent.mkdir(parents=True, exist_ok=True)
    template.save(str(out_path))
    return {"out": str(out_path), "reports": [report.__dict__ for report in reports], "counters": counters}
