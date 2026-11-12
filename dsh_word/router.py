"""Router layer: corrections in, batched fill + check tasks out.

The Router owns three decisions and nothing else:

* **batching** - corrections that touch the same slot kind travel together, so a
  burst of five "image too small" messages becomes one fill task;
* **patching** - a correction only overrides the slot it names; everything else
  keeps the base plan, which is what keeps "copy, never rewrite" true;
* **reporting** - each batch produces an output document plus the gate findings,
  which is the only thing the Router needs to decide the next step.

Nothing here talks to a model.  A model may decide *which* slot a user sentence
refers to; the mapping lands in this file as data.
"""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from docx import Document

from .checks import Finding, run_checks
from .fill import FillPlan, fill_docx, load_plan
from .media import image_extent, parse_length
from .omml import extract_math
from .values import FormulaSpec, ImageSpec, TableSpec

KINDS = {"text", "image", "formula", "table"}


class RouterError(RuntimeError):
    """A correction that cannot be applied without guessing."""


@dataclass
class Correction:
    slot: str
    changes: dict = field(default_factory=dict)
    note: str = ""

    @property
    def kind(self) -> str:
        return self.slot.split(".", 1)[0]


@dataclass
class Batch:
    kind: str
    corrections: list[Correction] = field(default_factory=list)


@dataclass
class BatchResult:
    batch: Batch
    out_path: str
    operations: list[dict] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)
    remaining: list[dict] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "kind": self.batch.kind,
            "slots": [correction.slot for correction in self.batch.corrections],
            "output": self.out_path,
            "operations": len(self.operations),
            "findings": [finding.as_dict() for finding in self.findings],
            "remaining": self.remaining,
        }

    @property
    def ok(self) -> bool:
        return not self.remaining and not any(
            finding.severity == "error" for finding in self.findings
        )


def load_corrections(path: str | Path) -> list[Correction]:
    """Read a correction queue; later entries for a slot win, notes are kept."""
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    items = data.get("corrections") if isinstance(data, dict) else data
    if not isinstance(items, list):
        raise RouterError("corrections file must contain a list under 'corrections'")

    corrections: list[Correction] = []
    for item in items:
        if not isinstance(item, dict) or "slot" not in item:
            raise RouterError(f"each correction needs a 'slot': {item!r}")
        slot = str(item["slot"])
        changes = {
            key: value
            for key, value in item.items()
            if key not in {"slot", "note"}
        }
        corrections.append(Correction(slot, changes, str(item.get("note", ""))))
    return corrections


def _validate(correction: Correction) -> None:
    """Reject incomplete corrections before any document work happens."""
    kind = correction.kind
    if kind not in KINDS:
        raise RouterError(
            f"slot {correction.slot!r} has unknown kind {kind!r}; "
            f"expected one of {sorted(KINDS)}"
        )
    changes = correction.changes
    if kind == "text" and not isinstance(changes.get("value"), str):
        raise RouterError(f"text correction for {correction.slot!r} needs a string 'value'")
    if kind == "image" and not ({"path", "width", "height"} & set(changes)):
        raise RouterError(
            f"image correction for {correction.slot!r} needs 'path', 'width' or 'height'"
        )
    if kind == "formula" and ("doc" not in changes or "paragraph" not in changes):
        raise RouterError(
            f"formula correction for {correction.slot!r} needs 'doc' and 'paragraph'"
        )
    if kind == "table" and ("doc" not in changes or "table" not in changes):
        raise RouterError(
            f"table correction for {correction.slot!r} needs 'doc' and 'table'"
        )


def batch_corrections(corrections: list[Correction]) -> list[Batch]:
    """Group by slot kind, merge repeated slots (last change set wins)."""
    merged: dict[str, Correction] = {}
    order: list[str] = []
    for correction in corrections:
        _validate(correction)
        if correction.slot in merged:
            previous = merged[correction.slot]
            changes = {**previous.changes, **correction.changes}
            note = " / ".join(filter(None, [previous.note, correction.note]))
            merged[correction.slot] = Correction(correction.slot, changes, note)
        else:
            merged[correction.slot] = correction
            order.append(correction.slot)

    batches: dict[str, Batch] = {}
    for slot in order:
        correction = merged[slot]
        batches.setdefault(correction.kind, Batch(correction.kind))
        batches[correction.kind].corrections.append(correction)
    return list(batches.values())


def _patch_text(plan: FillPlan, correction: Correction, base_dir: Path) -> None:
    value = correction.changes.get("value")
    if not isinstance(value, str):
        raise RouterError(f"text correction for {correction.slot!r} needs a string 'value'")
    plan.text[correction.slot] = value


def _patch_image(plan: FillPlan, correction: Correction, base_dir: Path) -> None:
    existing = plan.images.get(correction.slot)
    raw_path = correction.changes.get("path")
    path = (base_dir / str(raw_path)).resolve() if raw_path else existing and existing.path
    if not path:
        raise RouterError(f"image correction for {correction.slot!r} needs 'path'")
    width = (
        parse_length(correction.changes["width"])
        if "width" in correction.changes
        else existing and existing.width
    )
    height = (
        parse_length(correction.changes["height"])
        if "height" in correction.changes
        else existing and existing.height
    )
    cx, cy = image_extent(path, width, height)
    plan.images[correction.slot] = ImageSpec(str(path), cx, cy)


def _patch_formula(plan: FillPlan, correction: Correction, base_dir: Path) -> None:
    changes = correction.changes
    if "doc" not in changes or "paragraph" not in changes:
        raise RouterError(
            f"formula correction for {correction.slot!r} needs 'doc' and 'paragraph'"
        )
    doc = Document(str((base_dir / str(changes["doc"])).resolve()))
    element = extract_math(doc.paragraphs[int(changes["paragraph"])])
    plan.formulas[correction.slot] = FormulaSpec(
        element, display=bool(changes.get("display", True))
    )


def _patch_table(plan: FillPlan, correction: Correction, base_dir: Path) -> None:
    changes = correction.changes
    if "doc" not in changes or "table" not in changes:
        raise RouterError(
            f"table correction for {correction.slot!r} needs 'doc' and 'table'"
        )
    doc = Document(str((base_dir / str(changes["doc"])).resolve()))
    table = doc.tables[int(changes["table"])]
    rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
    if changes.get("skip_header"):
        rows = rows[1:]
    plan.tables[correction.slot] = TableSpec(rows)


_PATCHERS = {
    "text": _patch_text,
    "image": _patch_image,
    "formula": _patch_formula,
    "table": _patch_table,
}


def apply_batch(
    batch: Batch,
    template_path: str | Path,
    plan: FillPlan,
    out_path: str | Path,
    base_dir: str | Path | None = None,
) -> BatchResult:
    """Patch a copy of the base plan, refill the template, then run the gates."""
    patched = copy.deepcopy(plan)
    base = Path(base_dir) if base_dir else Path.cwd()
    for correction in batch.corrections:
        _PATCHERS[batch.kind](patched, correction, base)

    out_path = Path(out_path)
    fill = fill_docx(template_path, patched, out_path)
    findings = run_checks(out_path, template_path=template_path)
    return BatchResult(batch, str(out_path), fill.trace, findings, fill.remaining)


def run_router(
    template_path: str | Path,
    base_plan_path: str | Path,
    corrections_path: str | Path,
    out_dir: str | Path,
    report_path: str | Path | None = None,
    dry_run: bool = False,
) -> dict:
    """Batch the queue, apply each batch independently, and summarise results."""
    template_path = Path(template_path)
    base_plan_path = Path(base_plan_path)
    corrections_path = Path(corrections_path)
    out_dir = Path(out_dir)
    corrections = load_corrections(corrections_path)
    batches = batch_corrections(corrections)

    summary = {
        "template": str(template_path),
        "base_plan": str(base_plan_path),
        "corrections": len(corrections),
        "batches": [
            {
                "kind": batch.kind,
                "slots": [correction.slot for correction in batch.corrections],
                "notes": [correction.note for correction in batch.corrections if correction.note],
            }
            for batch in batches
        ],
        "dry_run": dry_run,
        "results": [],
    }
    if dry_run:
        return summary

    plan = load_plan(base_plan_path)
    out_dir.mkdir(parents=True, exist_ok=True)
    results: list[BatchResult] = []
    for index, batch in enumerate(batches, start=1):
        out_path = out_dir / f"batch{index:02d}-{batch.kind}.docx"
        result = apply_batch(
            batch,
            template_path,
            plan,
            out_path,
            base_dir=corrections_path.parent,
        )
        results.append(result)
        summary["results"].append(result.as_dict())

    summary["ok"] = all(result.ok for result in results)
    if report_path is not None:
        report_path = Path(report_path)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    return summary
