"""Slot replacement inside a paragraph, preserving the template's run formatting.

Markers look like ``{{ text.intro }}`` and may be split across several runs (Word
does this whenever the user edits mid-word).  The replacement therefore works on
character offsets and writes into the run that already carries the template
formatting, so fonts, sizes, east-asian fonts and bold settings survive untouched.

Field codes (``w:instrText``) are never seen here: python-docx only exposes
``w:t`` text on a run, which is exactly what protects Zotero fields.
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass

from docx.shared import Emu

from .values import FormulaSpec, ImageSpec

MARKER_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_.\-]*)\s*\}\}")


class SlotError(RuntimeError):
    """A slot could not be applied without damaging the paragraph."""


@dataclass(frozen=True)
class Marker:
    key: str
    start: int
    end: int


def find_markers(paragraph) -> list[Marker]:
    text = "".join(run.text for run in paragraph.runs)
    return [
        Marker(match.group(1), match.start(), match.end())
        for match in MARKER_RE.finditer(text)
    ]


def marker_keys(paragraph) -> list[str]:
    return [marker.key for marker in find_markers(paragraph)]


def _offsets(paragraph) -> list[tuple[object, int, int]]:
    offsets = []
    position = 0
    for run in paragraph.runs:
        text = run.text
        offsets.append((run, position, position + len(text)))
        position += len(text)
    return offsets


def _covered(offsets, start: int, end: int):
    covered = [
        (run, run_start, run_end)
        for run, run_start, run_end in offsets
        if run_end > start and run_start < end
    ]
    if not covered:
        raise SlotError(f"no run covers characters {start}:{end}")
    return covered


def _drop_if_empty(run) -> None:
    element = run._r
    if not run.text and element.getparent() is not None:
        element.getparent().remove(element)


def replace_text(paragraph, marker: Marker, value: str) -> None:
    """Substitute text for the marker, keeping the first covered run's format."""
    covered = _covered(_offsets(paragraph), marker.start, marker.end)
    first, first_start, _ = covered[0]
    first_text = first.text
    local = max(0, min(len(first_text), marker.start - first_start))
    head = first_text[:local]

    if len(covered) == 1:
        tail = first_text[marker.end - first_start:]
        first.text = head + value + tail
        return

    first.text = head + value
    last, last_start, last_end = covered[-1]
    cut = max(0, min(len(last.text), marker.end - last_start))
    last.text = last.text[cut:]
    for run, _, _ in covered[1:-1]:
        run.text = ""
        _drop_if_empty(run)
    if not last.text and last is not first:
        _drop_if_empty(last)


def replace_image(paragraph, marker: Marker, spec: ImageSpec) -> None:
    """Insert the picture where the marker sat, at the template's size."""
    covered = _covered(_offsets(paragraph), marker.start, marker.end)
    anchor = covered[0][0]
    width = Emu(spec.width) if spec.width else None
    height = Emu(spec.height) if spec.height else None

    run = paragraph.add_run()
    run.add_picture(spec.path, width=width, height=height)
    anchor._r.addprevious(run._r)

    for item, _, _ in covered:
        item.text = ""
    for item, _, _ in covered:
        _drop_if_empty(item)


def replace_formula(paragraph, marker: Marker, spec: FormulaSpec) -> None:
    """Deep-copy the OMML node in place; the equation stays editable."""
    covered = _covered(_offsets(paragraph), marker.start, marker.end)
    anchor = covered[0][0]._r
    anchor.addprevious(copy.deepcopy(spec.element))
    for item, _, _ in covered:
        item.text = ""
    for item, _, _ in covered:
        _drop_if_empty(item)


def replace_any(paragraph, marker: Marker, value) -> str:
    """Dispatch on the value type; returns the kind actually written."""
    if isinstance(value, str):
        replace_text(paragraph, marker, value)
        return "text"
    if isinstance(value, ImageSpec):
        replace_image(paragraph, marker, value)
        return "image"
    if isinstance(value, FormulaSpec):
        replace_formula(paragraph, marker, value)
        return "formula"
    raise SlotError(f"unsupported slot value for {marker.key!r}: {type(value).__name__}")
