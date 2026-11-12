"""Structure analysis, LLM-plan execution and verification for template mapping.

Pipeline:  analyze(template) + analyze(source)  →  mapping JSON  →  verify  →
apply.  The mapping is produced by the model in the Host, or deterministically
ordered when no model is available; this module only executes and checks it.
"""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu
from docx.text.paragraph import Paragraph

from .autoplan import _has_image, _save_images
from .fill import _save_highlight_copy
from .media import image_extent
from .omml import extract_math, math_nodes
from .smartfill import _append_like, _clear_runs, _set_formula, _set_image, _set_text


def _kind(paragraph: Paragraph) -> str:
    style = paragraph.style.name if paragraph.style is not None else "Normal"
    if style.lower().startswith(("heading", "标题", "title")):
        return "heading"
    if math_nodes(paragraph):
        return "formula"
    if _has_image(paragraph):
        return "image"
    if not paragraph.text.strip():
        return "empty"
    return "paragraph"


def analyze(path: str | Path) -> dict:
    document = Document(str(path))
    blocks: list[dict] = []
    sections: list[dict] = []
    paragraph_index = 0
    table_index = 0
    section = ""
    for child in document.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            paragraph = Paragraph(child, document)
            kind = _kind(paragraph)
            text = paragraph.text.strip()
            block = {
                "id": f"{kind}:{paragraph_index}",
                "kind": kind,
                "index": paragraph_index,
                "style": paragraph.style.name if paragraph.style is not None else "Normal",
                "chars": len(text),
                "text": text[:4000],
                "section": section,
            }
            blocks.append(block)
            if kind == "heading":
                section = text
                sections.append({"heading": text, "headingId": block["id"], "blocks": []})
            elif sections:
                sections[-1]["blocks"].append(block["id"])
            paragraph_index += 1
        elif tag == "tbl":
            table = document.tables[table_index]
            header = " | ".join(cell.text.strip() for cell in table.rows[0].cells) if table.rows else ""
            block = {
                "id": f"table:{table_index}",
                "kind": "table",
                "index": table_index,
                "rows": len(table.rows),
                "cols": len(table.columns),
                "text": header[:200],
                "section": section,
            }
            blocks.append(block)
            if sections:
                sections[-1]["blocks"].append(block["id"])
            table_index += 1
    return {"path": str(path), "blocks": blocks, "sections": sections}


@dataclass
class ApplyResult:
    out_path: str
    applied: int = 0
    skipped: list[str] = field(default_factory=list)
    appended: int = 0
    mapping: list[dict] = field(default_factory=list)


def _paragraphs(document) -> list[Paragraph]:
    return list(document.paragraphs)


def _fit_width(document, width: int) -> int:
    try:
        section = document.sections[0]
        available = section.page_width - section.left_margin - section.right_margin
        if available and width > available:
            return int(available)
    except Exception:  # noqa: BLE001 - geometry is best effort
        pass
    return width


def _block_issues(kind: str, paragraph: Paragraph, table_element=None) -> list[str]:
    """Per-block check: run right after a block is written."""
    issues: list[str] = []
    if kind in {"paragraph", "empty"}:
        text = paragraph.text
        if not text.strip():
            issues.append("文本为空")
        if "{{" in text:
            issues.append("残留占位符")
        if "\ufffd" in text:
            issues.append("含替换字符")
    elif kind == "image":
        if not _has_image(paragraph):
            issues.append("图片未插入")
    elif kind == "formula":
        if not math_nodes(paragraph):
            issues.append("公式未写入")
    elif kind == "table":
        rows = 0
        if table_element is not None:
            rows = len(table_element.findall(".//" + qn("w:tr")))
        if rows == 0:
            issues.append("表格为空")
    return issues


def apply_mapping(
    template_path: str | Path,
    source_path: str | Path,
    mapping: list[dict],
    out_path: str | Path,
    steps_dir: str | Path | None = None,
    on_op=None,
    on_check=None,
) -> ApplyResult:
    template = Document(str(template_path))
    source = Document(str(source_path))
    result = ApplyResult(out_path=str(out_path))
    assets = Path(out_path).parent / "assets"
    steps_root = Path(steps_dir) if steps_dir else None
    if steps_root is not None:
        steps_root.mkdir(parents=True, exist_ok=True)
    reference = template.paragraphs[-1] if template.paragraphs else None

    def emit(index: int, kind: str, detail: str) -> None:
        if steps_root is not None:
            step_path = steps_root / f"{Path(out_path).stem}.step.{index:02d}.docx"
            template.save(str(step_path))
            _save_highlight_copy(
                str(step_path),
                steps_root,
                Path(out_path).stem,
                index,
                {"kind": "body", "paragraph": len(template.paragraphs) - 1},
            )
        if on_op is not None:
            on_op(
                {
                    "op": index,
                    "kind": kind,
                    "slot": f"{kind}.m{index}",
                    "location": {"kind": "body", "paragraph": len(template.paragraphs) - 1},
                    "detail": detail,
                    "step": str(steps_root / f"{Path(out_path).stem}.step.{index:02d}.docx")
                    if steps_root
                    else "",
                }
            )

    anchor = None

    def insert_paragraph_after(element, donor):
        paragraph = template.add_paragraph()
        if donor is not None and donor._p.pPr is not None:
            paragraph._p.insert(0, copy.deepcopy(donor._p.pPr))
        if element is not None:
            element.addnext(paragraph._p)
        return paragraph

    for index, item in enumerate(mapping, start=1):
        action = item.get("action", "replace")
        template_id = str(item.get("template", ""))
        source_id = str(item.get("source", ""))
        if action == "skip" or not source_id:
            result.skipped.append(template_id)
            continue
        source_kind, _, source_index_text = source_id.partition(":")
        source_index = int(source_index_text) if source_index_text else 0
        template_kind, _, template_index_text = template_id.partition(":")
        template_index = int(template_index_text) if template_index_text else 0
        paragraphs = _paragraphs(template)

        donor = None
        if template_kind in {"paragraph", "empty"} and template_index < len(paragraphs):
            donor = paragraphs[template_index]

        if source_kind == "table":
            if source_index >= len(source.tables):
                result.skipped.append(source_id)
                continue
            element = copy.deepcopy(source.tables[source_index]._tbl)
            if template_kind == "table" and template_index < len(template.tables):
                target = template.tables[template_index]._tbl
                target.addnext(element)
                target.getparent().remove(target)
            elif anchor is not None:
                anchor.addnext(element)
            else:
                template.element.body.append(element)
            anchor = element
            result.applied += 1
            emit(index, "table", f"表格 {source_index + 1}")
            if on_check is not None:
                block_issues = _block_issues("table", donor, element)
                on_check(
                    {
                        "index": index,
                        "template": template_id,
                        "source": source_id,
                        "kind": "table",
                        "ok": not block_issues,
                        "issues": block_issues,
                    }
                )
            continue

        if action in {"insert-after", "append"} or template_kind not in {"paragraph", "empty"}:
            if anchor is not None:
                paragraph = insert_paragraph_after(anchor, donor)
            else:
                paragraph = _append_like(template, reference)
            anchor = paragraph._p
            result.appended += 1
        else:
            paragraph = paragraphs[template_index]
            anchor = paragraph._p

        source_paragraph = source.paragraphs[source_index]
        if source_kind == "image" or _has_image(source_paragraph):
            saved = _save_images(source_paragraph, source, assets, index)
            if not saved:
                result.skipped.append(source_id)
                continue
            width, _ = image_extent(str(saved[0]))
            _set_image(paragraph, str(saved[0]), _fit_width(template, width))
            detail = saved[0].name
        elif source_kind == "formula" or math_nodes(source_paragraph):
            _set_formula(paragraph, extract_math(source_paragraph))
            detail = f"公式 {index}"
        else:
            _set_text(paragraph, source_paragraph.text.strip())
            detail = source_paragraph.text.strip()[:60]

        result.applied += 1
        result.mapping.append({"template": template_id, "source": source_id, "action": action})
        emit(index, source_kind, detail)
        if on_check is not None:
            block_issues = _block_issues(source_kind, paragraph)
            on_check(
                {
                    "index": index,
                    "template": template_id,
                    "source": source_id,
                    "kind": source_kind,
                    "ok": not block_issues,
                    "issues": block_issues,
                }
            )

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    template.save(str(out_path))
    return result


def verify_mapping(template: dict, source: dict, mapping: list[dict]) -> dict:
    issues: list[dict] = []
    template_ids = {block["id"]: block for block in template["blocks"]}
    source_ids = {block["id"]: block for block in source["blocks"]}
    used_sources: set[str] = set()
    for item in mapping:
        template_id = str(item.get("template", ""))
        source_id = str(item.get("source", ""))
        action = item.get("action", "replace")
        if action == "skip" or not source_id:
            continue
        if template_id == "append":
            used_sources.add(source_id)
            continue
        if template_id not in template_ids:
            issues.append({"level": "error", "message": f"模板块不存在：{template_id}"})
        if source_id not in source_ids:
            issues.append({"level": "error", "message": f"资料块不存在：{source_id}"})
            continue
        if source_id in used_sources:
            issues.append({"level": "error", "message": f"资料块被重复使用：{source_id}"})
        used_sources.add(source_id)
        source_kind = source_ids[source_id]["kind"]
        template_kind = template_ids.get(template_id, {}).get("kind")
        if template_kind == "table" and source_kind != "table":
            issues.append({"level": "error", "message": f"表格只能放表格：{template_id} ← {source_id}"})
        if (
            template_kind in {"paragraph", "empty"}
            and source_kind == "table"
            and action == "replace"
        ):
            issues.append({"level": "warn", "message": f"表格替换了正文段落：{template_id} ← {source_id}"})
        if source_kind == "paragraph":
            chars = source_ids[source_id].get("chars", 0)
            if chars > 1200:
                issues.append({"level": "warn", "message": f"文本过长（{chars} 字）：{source_id}"})

    for block in source["blocks"]:
        if block["id"] in used_sources:
            continue
        if block["kind"] == "heading":
            issues.append({"level": "info", "message": f"资料标题未映射：{block['id']}"})
        else:
            issues.append(
                {"level": "error", "message": f"资料内容未被映射：{block['id']}（{block['kind']}）"}
            )
    body_blocks = [
        block for block in template["blocks"] if block["kind"] in {"paragraph", "empty"}
    ]
    mapped = {item.get("template") for item in mapping}
    for block in body_blocks:
        if block["id"] not in mapped:
            issues.append({"level": "info", "message": f"模板段落未映射：{block['id']}"})
    return {"ok": not any(issue["level"] == "error" for issue in issues), "issues": issues}
