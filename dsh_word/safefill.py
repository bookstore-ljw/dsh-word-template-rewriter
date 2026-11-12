"""Non-destructive fill: change paragraph text only, never the template layout.

The template file is copied verbatim; only the text inside existing paragraphs of
recognised sections is rewritten.  No paragraph is added, removed, or re-styled,
so single-column and two-column templates keep their formatting.  Content that
does not fit existing placeholders is reported instead of forcing layout changes.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document

from .smartfill import _clear_runs, _set_text
from .sectional import _is_heading, label_of, segment, _text_of_source_blocks


def _section_lines(label: str, content: dict, body_state: dict) -> list[str]:
    if label == "title":
        return [content["title"]] if content["title"] else []
    if label == "authors":
        return list(content["authors"])
    if label == "abstract":
        return [content["abstract"]] if content["abstract"] else []
    if label == "keywords":
        return [content["keywords"]] if content["keywords"] else []
    body = content["body"]
    index = body_state["index"]
    if index < len(body):
        body_state["index"] = index + 1
        return list(body[index][1])
    return []


def safe_fill(template_path, source_path, out_path, steps_dir=None, on_event=None):
    template_path = Path(template_path)
    source_path = Path(source_path)
    out_path = Path(out_path)
    document = Document(str(template_path))
    before_count = len(document.paragraphs)
    source = Document(str(source_path))
    content = _text_of_source_blocks(source)
    sections = segment(document)
    body_state = {"index": 0}
    reports: list[dict] = []
    leftover: list[str] = []

    for section in sections:
        label = section["label"]
        lines = _section_lines(label, content, body_state)
        paragraphs = document.paragraphs
        targets = [
            index
            for index in range(section["start"] + 1, min(section["end"], len(paragraphs)))
            if not _is_heading(paragraphs[index])
        ]
        replaced = 0
        for offset, index in enumerate(targets):
            paragraph = paragraphs[index]
            if offset < len(lines):
                _set_text(paragraph, lines[offset])
                replaced += 1
            # extra template paragraphs are left untouched: clearing them leaves
            # blank lines and section-break artefacts, which looks like broken layout
        if len(lines) > len(targets):
            leftover.extend(lines[len(targets):])
        reports.append(
            {
                "label": label,
                "title": section.get("title", ""),
                "replaced": replaced,
                "slots": len(targets),
                "leftover": max(0, len(lines) - len(targets)),
            }
        )
        if steps_dir is not None:
            step_dir = Path(steps_dir)
            step_dir.mkdir(parents=True, exist_ok=True)
            document.save(str(step_dir / f"{out_path.stem}.step.{len(reports):02d}.docx"))
        if on_event:
            on_event(reports[-1])

    after_count = len(document.paragraphs)
    if after_count != before_count:
        raise RuntimeError(
            f"非破坏填充异常：段落数从 {before_count} 变成 {after_count}（不应增删段落）"
        )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(out_path))
    return {
        "out": str(out_path),
        "reports": reports,
        "leftover": leftover,
        "media_skipped": True,
    }
