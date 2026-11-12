"""Restore template formatting on paragraphs that drifted during filling.

The fill keeps paragraph properties and run formats, but a replaced paragraph can
still lose its style (for example ``Heading 2`` becomes ``Normal``).  This pass
copies the template's ``pPr``/run font back onto the matching paragraph and
restores the page setup, so the finished document matches the template again.
"""

from __future__ import annotations

import copy
from pathlib import Path

from docx import Document


def heal_format(template_path: str | Path, docx_path: str | Path) -> dict:
    template = Document(str(template_path))
    document = Document(str(docx_path))
    healed_paragraphs = 0
    healed_runs = 0
    healed_pages = 0

    template_paragraphs = list(template.paragraphs)
    document_paragraphs = list(document.paragraphs)
    for before, after in zip(template_paragraphs, document_paragraphs):
        changed = False
        if before._p.pPr is not None:
            existing = after._p.pPr
            if existing is not None:
                after._p.remove(existing)
            after._p.insert(0, copy.deepcopy(before._p.pPr))
            changed = True
        if (before.style.name or "") != (after.style.name or ""):
            try:
                after.style = document.styles[before.style.name]
                changed = True
            except KeyError:
                pass
        reference = before.runs[0].font if before.runs else None
        if reference is not None:
            for run in after.runs:
                if reference.name and run.font.name != reference.name:
                    run.font.name = reference.name
                    healed_runs += 1
                if reference.size and run.font.size != reference.size:
                    run.font.size = reference.size
                    healed_runs += 1
        if changed:
            healed_paragraphs += 1

    if template.sections and document.sections:
        before = template.sections[0]
        after = document.sections[0]
        for name in (
            "left_margin",
            "right_margin",
            "top_margin",
            "bottom_margin",
            "page_width",
            "page_height",
            "orientation",
        ):
            value = getattr(before, name, None)
            if value is not None and getattr(after, name, None) != value:
                setattr(after, name, value)
                healed_pages += 1

    document.save(str(docx_path))
    return {
        "paragraphs": healed_paragraphs,
        "runs": healed_runs,
        "page": healed_pages,
    }


def normalize_headings(docx_path: str | Path) -> dict:
    """Clamp heading levels so they never jump by more than one."""
    document = Document(str(docx_path))
    previous = 0
    changed = 0
    for paragraph in document.paragraphs:
        style = paragraph.style.name if paragraph.style is not None else ""
        lowered = style.lower()
        if not (lowered.startswith(("heading", "标题"))):
            continue
        digits = [c for c in style if c.isdigit()]
        level = int(digits[0]) if digits else 1
        if level > previous + 1:
            target = f"Heading {previous + 1}"
            try:
                paragraph.style = document.styles[target]
                changed += 1
            except KeyError:
                try:
                    paragraph.style = document.styles[f"标题 {previous + 1}"]
                    changed += 1
                except KeyError:
                    pass
            level = previous + 1
        previous = level
    if changed:
        document.save(str(docx_path))
    return {"normalized": changed}
