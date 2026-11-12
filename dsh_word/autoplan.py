"""Automatic slot mapping for the panel: template + source in, plan out.

The mapping is deliberately conservative and inspectable: it never invents text,
it only decides *which* source element feeds which template slot, and reports
every decision.  Missing material becomes an explicit unresolved entry instead of
a silently empty slot.

Order is the only heuristic:

* ``text.*``    - body paragraphs that carry text, in document order
* ``formula.*`` - paragraphs that carry OMML, in document order
* ``image.*``   - embedded images, in document order (written next to the output)
* ``table.*``   - tables, in document order
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from .fill import FillPlan
from .media import image_extent
from .omml import extract_math, math_nodes
from .ooxml import iter_paragraphs
from .slots import find_markers
from .values import FormulaSpec, ImageSpec, TableSpec

_BLIP = qn("a:blip")
_EMBED = qn("r:embed")


@dataclass
class AutoPlan:
    plan: FillPlan
    mapping: list[dict] = field(default_factory=list)
    unresolved: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {"mapping": self.mapping, "unresolved": self.unresolved}


def _template_slots(template) -> list[tuple[str, str]]:
    """Every marker in document order, with a rough location label."""
    slots: list[tuple[str, str]] = []
    for location, paragraph in iter_paragraphs(template):
        for marker in find_markers(paragraph):
            slots.append((marker.key, location.label()))
    return slots


def _has_image(paragraph) -> bool:
    return bool(paragraph._p.findall(".//" + _BLIP))


def _save_images(paragraph, document, asset_dir: Path, index: int) -> list[Path]:
    saved: list[Path] = []
    for blip in paragraph._p.findall(".//" + _BLIP):
        relationship_id = blip.get(_EMBED)
        if not relationship_id:
            continue
        part = document.part.related_parts[relationship_id]
        suffix = Path(str(part.partname)).suffix or ".png"
        asset_dir.mkdir(parents=True, exist_ok=True)
        target = asset_dir / f"source-image-{index}{suffix}"
        target.write_bytes(part.blob)
        saved.append(target)
        index += 1
    return saved


def build_auto_plan(
    template_path: str | Path,
    source_path: str | Path,
    asset_dir: str | Path | None = None,
) -> AutoPlan:
    template = Document(str(template_path))
    source = Document(str(source_path))
    asset_root = Path(asset_dir) if asset_dir else Path(template_path).parent / "assets"

    slots = _template_slots(template)
    plan = FillPlan()
    auto = AutoPlan(plan)

    text_pool = []
    formula_pool = []
    image_pool: list[Path] = []
    for index, paragraph in enumerate(source.paragraphs):
        if math_nodes(paragraph):
            formula_pool.append((paragraph, f"source.p[{index}]"))
        elif _has_image(paragraph):
            for saved in _save_images(paragraph, source, asset_root, len(image_pool) + 1):
                image_pool.append(saved)
        elif paragraph.text.strip():
            text_pool.append((paragraph.text, f"source.p[{index}]"))

    table_pool = [
        (
            [[cell.text.strip() for cell in row.cells] for row in table.rows],
            f"source.table[{index}]",
        )
        for index, table in enumerate(source.tables)
    ]

    for slot, where in slots:
        kind = slot.split(".", 1)[0]
        if kind == "text":
            if text_pool:
                value, origin = text_pool.pop(0)
                plan.text[slot] = value
                auto.mapping.append(
                    {"slot": slot, "from": origin, "into": where, "kind": "text"}
                )
            else:
                auto.unresolved.append(slot)
        elif kind == "formula":
            if formula_pool:
                paragraph, origin = formula_pool.pop(0)
                plan.formulas[slot] = FormulaSpec(extract_math(paragraph))
                auto.mapping.append(
                    {"slot": slot, "from": origin, "into": where, "kind": "formula"}
                )
            else:
                auto.unresolved.append(slot)
        elif kind == "image":
            if image_pool:
                image_path = image_pool.pop(0)
                width, height = image_extent(image_path)
                plan.images[slot] = ImageSpec(str(image_path), width, height)
                auto.mapping.append(
                    {
                        "slot": slot,
                        "from": image_path.name,
                        "into": where,
                        "kind": "image",
                    }
                )
            else:
                auto.unresolved.append(slot)
        elif kind == "table":
            if table_pool:
                rows, origin = table_pool.pop(0)
                plan.tables[slot] = TableSpec(rows)
                auto.mapping.append(
                    {"slot": slot, "from": origin, "into": where, "kind": "table"}
                )
            else:
                auto.unresolved.append(slot)
        else:
            auto.unresolved.append(slot)

    return auto
