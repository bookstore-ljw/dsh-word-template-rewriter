"""Section-by-section filling driven by the template's own semantics.

The template is a form: its headings/labels tell us what each part is for
("标题", "作者", "摘要", "关键词", "参考文献", ...).  We segment the template in
order, find the matching material in the source for each segment, and report the
segments that cannot be filled instead of silently skipping them.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from docx import Document

from .mapping import analyze, apply_mapping

_LABELS = [
    ("title", ("标题", "题目", "title")),
    ("authors", ("作者", "author")),
    ("affiliation", ("单位", "机构", "affiliation")),
    ("abstract", ("摘要", "abstract")),
    ("keywords", ("关键词", "关键字", "keywords")),
    ("references", ("参考文献", "references", "bibliography")),
]


def label_of(text: str) -> str | None:
    lowered = text.strip().lower()
    for name, words in _LABELS:
        if any(word in lowered for word in words):
            return name
    return None


def _is_heading(block: dict) -> bool:
    return block.get("kind") == "heading"


@dataclass
class Missing:
    segment: str
    reason: str = ""
    extra: dict = field(default_factory=dict)


def segment_template(blocks: list[dict]) -> list[dict]:
    segments: list[dict] = []
    current = {"label": "front", "title": "", "blocks": []}
    for block in blocks:
        if _is_heading(block):
            if current["blocks"]:
                segments.append(current)
            current = {
                "label": label_of(block["text"]) or "body",
                "title": block["text"],
                "blocks": [],
            }
            continue
        current["blocks"].append(block["id"])
    if current["blocks"]:
        segments.append(current)
    return segments


def extract_source(blocks: list[dict]) -> dict:
    texts = [block for block in blocks if block["kind"] in {"paragraph", "heading"}]
    found: dict[str, object] = {
        "title": None,
        "authors": None,
        "affiliation": None,
        "abstract": None,
        "keywords": None,
        "references": [],
        "body": [],
    }
    if texts:
        found["title"] = texts[0]
    in_references = False
    for block in texts[1:]:
        label = label_of(block["text"])
        in_references = in_references or label == "references"
        if in_references:
            found["references"].append(block)
        elif label == "abstract" and not found["abstract"]:
            found["abstract"] = block
        elif label == "keywords" and not found["keywords"]:
            found["keywords"] = block
        elif label == "authors" and not found["authors"]:
            found["authors"] = block
        elif label == "affiliation" and not found["affiliation"]:
            found["affiliation"] = block
        elif found["authors"] is None and block["id"] != found["title"]["id"] and len(block["text"]) < 80:
            found["authors"] = block
        else:
            found["body"].append(block)
    return found


def build_section_mapping(template_blocks, source_blocks):
    segments = segment_template(template_blocks)
    material = extract_source(source_blocks)
    mapping: list[dict] = []
    missing: list[Missing] = []
    report: list[dict] = []

    queue = list(material["body"])
    for segment in segments:
        label = segment["label"]
        slots = [block_id for block_id in segment["blocks"]]
        if label == "front":
            picks = [item for item in (material["title"], material["authors"]) if item]
        elif label in {"title", "authors", "affiliation", "abstract", "keywords"}:
            picks = [material[label]] if material[label] else []
        elif label == "references":
            picks = list(material["references"])
        else:
            picks = queue
        if not picks:
            missing.append(Missing(label, "资料中没有对应内容", {"segment": segment.get("title", label)}))
            report.append({"segment": label, "title": segment.get("title", ""), "mapped": 0, "missing": True})
            continue
        count = 0
        for slot, source in zip(slots, picks):
            mapping.append({"template": slot, "source": source["id"], "action": "replace"})
            count += 1
            if label == "body" and source in queue:
                queue.remove(source)
        report.append({"segment": label, "title": segment.get("title", ""), "mapped": count, "missing": False})

    for block in queue:
        mapping.append({"template": "append", "source": block["id"], "action": "append"})
    return mapping, missing, report


def section_fill(template_path, source_path, out_path, steps_dir=None, on_event=None):
    template = analyze(template_path)
    source = analyze(source_path)
    mapping, missing, report = build_section_mapping(template["blocks"], source["blocks"])
    if on_event:
        for item in report:
            on_event({"type": "segment", **item})
        for item in missing:
            on_event({"type": "missing", "segment": item.segment, "reason": item.reason})
    result = apply_mapping(
        template_path,
        source_path,
        mapping,
        out_path,
        steps_dir=steps_dir,
        on_op=(lambda item: on_event({"type": "progress", **item})) if on_event else None,
        on_check=(lambda item: on_event({"type": "check", **item})) if on_event else None,
    )
    return {"result": result, "missing": [item.__dict__ for item in missing], "report": report}
