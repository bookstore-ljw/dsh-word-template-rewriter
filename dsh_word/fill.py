"""The fill worker: copy content into a template, one atomic operation at a time.

The worker never writes prose.  It resolves each slot to source material (a
paragraph, an image file, an OMML node, a table) and copies it into the place the
template reserved, then records what it did.  Every atomic operation can be saved
as an intermediate document so a preview pane can follow the work.
"""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import yaml
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Row as TableRow

from .media import image_extent, parse_length
from .ooxml import iter_paragraphs, paragraph_text
from .omml import extract_math, math_nodes, math_text
from .slots import Marker, find_markers, replace_any
from .values import FormulaSpec, ImageSpec, TableSpec

STEPS_TEMPLATE = "{stem}.step.{op:02d}.docx"


class FillError(RuntimeError):
    """The template and the plan disagree in a way that cannot be copied safely."""


@dataclass
class FillPlan:
    text: dict[str, str] = field(default_factory=dict)
    images: dict[str, ImageSpec] = field(default_factory=dict)
    formulas: dict[str, FormulaSpec] = field(default_factory=dict)
    tables: dict[str, TableSpec] = field(default_factory=dict)

    def value_for(self, key: str):
        if key in self.text:
            return self.text[key]
        if key in self.images:
            return self.images[key]
        if key in self.formulas:
            return self.formulas[key]
        return None

    def known_keys(self) -> set[str]:
        return set(self.text) | set(self.images) | set(self.formulas) | set(self.tables)


@dataclass
class FillResult:
    out_path: str
    steps: list[str] = field(default_factory=list)
    trace: list[dict] = field(default_factory=list)
    remaining: list[dict] = field(default_factory=list)


# --------------------------------------------------------------------------- plan


def _qualify(prefix: str, key: str) -> str:
    """Allow ``title:`` in the plan while the template says ``{{ text.title }}``."""
    return key if key.startswith(prefix + ".") else f"{prefix}.{key}"


def _source_document(base: Path, spec: dict):
    doc_path = (base / str(spec["doc"])).resolve()
    if not doc_path.exists():
        raise FillError(f"source document not found: {doc_path}")
    return Document(str(doc_path))


def _source_paragraph_text(base: Path, spec: dict) -> str:
    doc = _source_document(base, spec)
    index = int(spec["paragraph"])
    if index >= len(doc.paragraphs):
        raise FillError(f"source {spec['doc']} has no body paragraph {index}")
    text = doc.paragraphs[index].text
    if not text.strip():
        raise FillError(f"source {spec['doc']} paragraph {index} is empty")
    return text


def load_plan(path: str | Path) -> FillPlan:
    """Read a YAML plan and resolve it into concrete, copyable values."""
    plan_path = Path(path)
    data = yaml.safe_load(plan_path.read_text(encoding="utf-8")) or {}
    base = plan_path.parent
    plan = FillPlan()

    for key, spec in (data.get("text") or {}).items():
        slot = _qualify("text", key)
        if isinstance(spec, str):
            plan.text[slot] = spec
        elif isinstance(spec, dict) and "doc" in spec:
            plan.text[slot] = _source_paragraph_text(base, spec)
        else:
            raise FillError(f"text slot {key!r} must be a string or {{doc, paragraph}}")

    for key, spec in (data.get("images") or {}).items():
        slot = _qualify("image", key)
        if isinstance(spec, str):
            spec = {"path": spec}
        image_path = (base / str(spec["path"])).resolve()
        if not image_path.exists():
            raise FillError(f"image not found for slot {key!r}: {image_path}")
        width = parse_length(spec.get("width"))
        height = parse_length(spec.get("height"))
        cx, cy = image_extent(image_path, width, height)
        plan.images[slot] = ImageSpec(str(image_path), cx, cy)

    for key, spec in (data.get("formulas") or {}).items():
        slot = _qualify("formula", key)
        doc = _source_document(base, spec)
        index = int(spec["paragraph"])
        if index >= len(doc.paragraphs):
            raise FillError(f"source {spec['doc']} has no body paragraph {index}")
        element = extract_math(doc.paragraphs[index])
        plan.formulas[slot] = FormulaSpec(element, display=bool(spec.get("display", True)))

    for key, spec in (data.get("tables") or {}).items():
        slot = _qualify("table", key)
        doc = _source_document(base, spec)
        table_index = int(spec["table"])
        if table_index >= len(doc.tables):
            raise FillError(f"source {spec['doc']} has no table {table_index}")
        rows = [
            [cell.text.strip() for cell in row.cells]
            for row in doc.tables[table_index].rows
        ]
        if spec.get("skip_header"):
            rows = rows[1:]
        plan.tables[slot] = TableSpec(rows)

    return plan


# ------------------------------------------------------------------- table filling


def _set_cell_text(cell, text: str) -> None:
    """Write into the cell's first paragraph, keeping its run formatting."""
    paragraphs = cell.paragraphs
    first = paragraphs[0]
    for extra in paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    runs = first.runs
    if not runs:
        first.add_run(text)
        return
    runs[0].text = text
    for run in runs[1:]:
        run._r.getparent().remove(run._r)


def _row_slot(row) -> str | None:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            keys = [marker.key for marker in find_markers(paragraph)]
            if keys:
                return keys[0]
    return None


def _fill_table(table, template_row, spec: TableSpec, slot: str) -> int:
    column_count = len(template_row.cells)
    for row_values in spec.rows:
        if len(row_values) > column_count:
            raise FillError(
                f"table slot {slot!r}: source row has {len(row_values)} cells, "
                f"template row has {column_count}"
            )
    for row_values in spec.rows:
        new_tr = copy.deepcopy(template_row._tr)
        template_row._tr.addprevious(new_tr)
        new_row = TableRow(new_tr, table)
        for cell_index, cell in enumerate(new_row.cells):
            value = row_values[cell_index] if cell_index < len(row_values) else ""
            _set_cell_text(cell, value)
    template_row._tr.getparent().remove(template_row._tr)
    return len(spec.rows)


# ---------------------------------------------------------------------- main entry


def _save_step(doc, steps_dir: Path, stem: str, op: int) -> str:
    steps_dir.mkdir(parents=True, exist_ok=True)
    path = steps_dir / STEPS_TEMPLATE.format(stem=stem, op=op)
    doc.save(str(path))
    return str(path)


def _shade_paragraph(paragraph, fill: str = "FFF2B2") -> None:
    pPr = paragraph._p.get_or_add_pPr()
    shd = pPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        pPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def _highlight_location(document, location: dict) -> None:
    """Mark the element this operation just changed, for the preview frame."""
    kind = location.get("kind")
    if kind == "body":
        index = location.get("paragraph")
        if isinstance(index, int) and 0 <= index < len(document.paragraphs):
            _shade_paragraph(document.paragraphs[index])
    elif kind == "table":
        index = location.get("table")
        if isinstance(index, int) and 0 <= index < len(document.tables):
            table = document.tables[index]
            if table.rows:
                for cell in table.rows[0].cells:
                    for paragraph in cell.paragraphs:
                        _shade_paragraph(paragraph)


def _save_highlight_copy(step_path: str, steps_dir: Path, stem: str, op: int, location: dict) -> None:
    try:
        document = Document(step_path)
        _highlight_location(document, location)
        target = steps_dir / "highlight" / STEPS_TEMPLATE.format(stem=stem, op=op)
        target.parent.mkdir(parents=True, exist_ok=True)
        document.save(str(target))
    except Exception:  # noqa: BLE001 - a preview aid must never break the fill
        pass


def fill_docx(
    template_path: str | Path,
    plan: FillPlan,
    out_path: str | Path,
    steps_dir: str | Path | None = None,
    trace_path: str | Path | None = None,
    on_op: Callable[[dict], None] | None = None,
) -> FillResult:
    template_path = Path(template_path)
    out_path = Path(out_path)
    doc = Document(str(template_path))
    steps_root = Path(steps_dir) if steps_dir else None
    result = FillResult(out_path=str(out_path))
    op = 0

    def record(kind: str, slot: str, location: dict, detail: str) -> None:
        nonlocal op
        op += 1
        step_path = ""
        if steps_root is not None:
            step_path = _save_step(doc, steps_root, out_path.stem, op)
            _save_highlight_copy(
                step_path, steps_root, out_path.stem, op, location
            )
            result.steps.append(step_path)
        entry = {
            "op": op,
            "kind": kind,
            "slot": slot,
            "location": location,
            "detail": detail,
            "step": step_path,
        }
        result.trace.append(entry)
        if on_op is not None:
            on_op(entry)

    # 1) Tables first: their paragraphs are cloned, so they must not be confused
    #    with body paragraphs handed to the slot resolver.
    filled_tables: set[int] = set()
    for table_index, table in enumerate(doc.tables):
        for row in list(table.rows):
            slot = _row_slot(row)
            if slot is None or slot not in plan.tables or table_index in filled_tables:
                continue
            rows = _fill_table(table, row, plan.tables[slot], slot)
            filled_tables.add(table_index)
            record("table", slot, {"kind": "table", "table": table_index}, f"{rows} rows")

    # 2) Body paragraphs, markers from right to left so offsets stay valid.
    for paragraph_index, paragraph in enumerate(doc.paragraphs):
        for marker in reversed(find_markers(paragraph)):
            value = plan.value_for(marker.key)
            if value is None:
                continue
            kind = replace_any(paragraph, marker, value)
            if isinstance(value, str):
                detail = value
            elif isinstance(value, ImageSpec):
                detail = f"{value.path} ({value.width}x{value.height} EMU)"
            elif isinstance(value, FormulaSpec):
                detail = math_text(value.element)
            else:
                detail = paragraph_text(paragraph)
            record(
                kind,
                marker.key,
                {"kind": "body", "paragraph": paragraph_index},
                str(detail)[:120],
            )

    # 3) Anything still holding a marker is reported, never silently dropped.
    for location, paragraph in iter_paragraphs(doc):
        for marker in find_markers(paragraph):
            result.remaining.append(
                {"slot": marker.key, "location": location.as_dict()}
            )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    if trace_path is not None:
        trace_path = Path(trace_path)
        trace_path.parent.mkdir(parents=True, exist_ok=True)
        trace_path.write_text(
            json.dumps(
                {
                    "template": str(template_path),
                    "output": str(out_path),
                    "operations": result.trace,
                    "remaining": result.remaining,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
    return result


def inspect_docx(path: str | Path) -> dict:
    """Cheap structural summary, used by the CLI and by smoke tests."""
    doc = Document(str(path))
    paragraphs = list(iter_paragraphs(doc))
    body = [p for location, p in paragraphs if location.kind == "body"]
    return {
        "path": str(path),
        "paragraphs": len(body),
        "tables": len(doc.tables),
        "images": len(doc.inline_shapes),
        "math": sum(len(math_nodes(p)) for _, p in paragraphs),
        "markers": sum(len(find_markers(p)) for _, p in paragraphs),
        "text_length": len(paragraph_text_from_doc(doc)),
    }


def paragraph_text_from_doc(doc) -> str:
    return "\n".join(paragraph_text(p) for _, p in iter_paragraphs(doc))
