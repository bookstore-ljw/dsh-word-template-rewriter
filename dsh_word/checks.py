"""Read-only quality gates.

Rules are data (see ``rules.yaml``), the engine only dispatches.  A check never
edits the document: it returns findings with a location, and the Router decides
whether to dispatch a fix back to the fill worker.
"""

from __future__ import annotations

import json
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from docx import Document

from .ooxml import (
    body_paragraphs,
    is_drawing_run,
    iter_paragraphs,
    paragraph_signature,
    paragraph_text,
    run_signature,
)
from .omml import iter_math, math_text, problems
from .slots import MARKER_RE, find_markers

DEFAULT_RULES = Path(__file__).with_name("rules.yaml")

CAMEL_RE = re.compile(r"\b[a-z]{3,}(?:[A-Z][a-z]{3,})+\b")
# "word.Next" is a copied-together sentence; "text.unknown" is an identifier.
PUNCT_GLUE_RE = re.compile(r"\b[a-z]{2,}[.;][A-Z][a-z]{2,}|\b[a-z]{2,}[,;:][a-z]{2,}")
TOKEN_RE = re.compile(r"\b[A-Za-z]{6,}\b")

STOPWORDS = set(
    """
    about above after again against all also and another any are around because
    been before being below between both but came can could did does doing down
    during each even ever every for from further had has have having here how
    however into its itself just like made make many may more most much must
    neither not now off once only other our out over own same should since some
    such than that the their them then there these they this those through to
    too under until up upon very was were what when where which while who whom
    why will with within without would you your
    """.split()
)


@dataclass
class Finding:
    rule: str
    severity: str
    message: str
    location: dict
    snippet: str = ""
    detail: str = ""

    def as_dict(self) -> dict:
        return {
            "rule": self.rule,
            "severity": self.severity,
            "message": self.message,
            "location": self.location,
            "snippet": self.snippet,
            "detail": self.detail,
        }


def load_rules(path: str | Path | None = None) -> list[dict]:
    rules_path = Path(path) if path else DEFAULT_RULES
    data = yaml.safe_load(rules_path.read_text(encoding="utf-8")) or {}
    rules = data.get("rules") or []
    for rule in rules:
        if "id" not in rule or "kind" not in rule:
            raise ValueError(f"invalid rule entry: {rule!r}")
    return [rule for rule in rules if rule.get("enabled", True)]


def _snippet(text: str, start: int, end: int, width: int = 18) -> str:
    left = max(0, start - width)
    right = min(len(text), end + width)
    excerpt = text[left:right].replace("\n", " ")
    return f"…{excerpt}…" if left or right < len(text) else excerpt


def _rule_message(rule: dict, fallback: str) -> str:
    return str(rule.get("message") or fallback)


# --------------------------------------------------------------------- handlers


def _check_text_regex(doc, template, rule) -> list[Finding]:
    flags = re.IGNORECASE if rule.get("ignore_case") else 0
    pattern = re.compile(rule["pattern"], flags)
    include_math = bool(rule.get("include_math"))
    findings: list[Finding] = []
    for location, paragraph in iter_paragraphs(doc):
        text = paragraph_text(paragraph)
        last_end = -1
        for match in pattern.finditer(text):
            if match.start() < last_end:  # "{{" and "}}" of one marker are one defect
                continue
            last_end = match.end()
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "warning"),
                    _rule_message(rule, "text matches forbidden pattern"),
                    location.as_dict(),
                    _snippet(text, match.start(), match.end()),
                )
            )
        if include_math:
            from .omml import math_nodes

            for node in math_nodes(paragraph):
                math = math_text(node)
                for match in pattern.finditer(math):
                    findings.append(
                        Finding(
                            rule["id"],
                            rule.get("severity", "warning"),
                            _rule_message(rule, "math text matches forbidden pattern"),
                            location.as_dict(),
                            _snippet(math, match.start(), match.end()),
                        )
                    )
    return findings


def _check_leftover_slot(doc, template, rule) -> list[Finding]:
    """Both whole markers and orphan braces are template defects worth reporting."""
    findings: list[Finding] = []
    for location, paragraph in iter_paragraphs(doc):
        text = paragraph_text(paragraph)
        for match in MARKER_RE.finditer(text):
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "error"),
                    _rule_message(rule, "placeholder was not replaced"),
                    location.as_dict(),
                    _snippet(text, match.start(), match.end()),
                    detail=match.group(0),
                )
            )
        marker_span = set()
        for match in MARKER_RE.finditer(text):
            marker_span.update(range(match.start(), match.end()))
        for match in re.finditer(r"\{\{|\}\}", text):
            if match.start() in marker_span:
                continue
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "error"),
                    _rule_message(rule, "orphan template brace"),
                    location.as_dict(),
                    _snippet(text, match.start(), match.end()),
                    detail="孤立的花括号，疑似手工删改模板",
                )
            )
    return findings


def _check_math_integrity(doc, template, rule) -> list[Finding]:
    findings: list[Finding] = []
    for location, node in iter_math(doc):
        found = problems(node)
        if found:
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "error"),
                    _rule_message(rule, "formula structure is incomplete"),
                    location.as_dict(),
                    math_text(node)[:80],
                    detail="; ".join(found),
                )
            )
    return findings


def _split_candidate(token: str, corpus: Counter, min_part: int):
    for index in range(min_part, len(token) - min_part + 1):
        left, right = token[:index], token[index:]
        if left != token and right != token and corpus[left] and corpus[right]:
            return left, right
    return None


def _check_glued_words(doc, template, rule) -> list[Finding]:
    allow = {str(word).lower() for word in rule.get("allow", [])}
    min_part = int(rule.get("min_part_length", 3))
    corpus: Counter = Counter()
    paragraphs = list(iter_paragraphs(doc))
    for _, paragraph in paragraphs:
        corpus.update(re.findall(r"[a-z]+", paragraph_text(paragraph).lower()))

    findings: list[Finding] = []
    for location, paragraph in paragraphs:
        text = paragraph_text(paragraph)
        for match in CAMEL_RE.finditer(text):
            token = match.group(0)
            if token.lower() in allow:
                continue
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "warning"),
                    _rule_message(rule, "suspicious concatenation"),
                    location.as_dict(),
                    _snippet(text, match.start(), match.end()),
                    detail=f"{token}: 小写后直接跟大写，疑似两个词粘连",
                )
            )
        for match in PUNCT_GLUE_RE.finditer(text):
            token = match.group(0)
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "warning"),
                    _rule_message(rule, "missing space after punctuation"),
                    location.as_dict(),
                    _snippet(text, match.start(), match.end()),
                    detail=f"{token}: 标点后紧跟单词，疑似缺少空格",
                )
            )
        for match in TOKEN_RE.finditer(text):
            token = match.group(0).lower()
            if token in allow:
                continue
            split = _split_candidate(token, corpus, min_part)
            if split is None:
                continue
            findings.append(
                Finding(
                    rule["id"],
                    "info",
                    _rule_message(rule, "possible glued words"),
                    location.as_dict(),
                    _snippet(text, match.start(), match.end()),
                    detail=f"{match.group(0)} ≈ {split[0]} {split[1]}（待确认）",
                )
            )
    return findings


def _check_conformance(doc, template, rule) -> list[Finding]:
    if template is None:
        return []
    findings: list[Finding] = []
    template_paragraphs = body_paragraphs(template)
    output_paragraphs = body_paragraphs(doc)

    if len(template_paragraphs) != len(output_paragraphs):
        findings.append(
            Finding(
                rule["id"],
                rule.get("severity", "error"),
                "正文段落数与模板不一致",
                {"kind": "body"},
                detail=f"template={len(template_paragraphs)} output={len(output_paragraphs)}",
            )
        )
        return findings

    for index, (expected, produced) in enumerate(zip(template_paragraphs, output_paragraphs)):
        if not find_markers(expected):
            continue
        location = {"kind": "body", "paragraph": index}
        expected_ppr = paragraph_signature(expected)
        produced_ppr = paragraph_signature(produced)
        if expected_ppr != produced_ppr:
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "error"),
                    "段落格式被改动",
                    location,
                    detail=f"{expected_ppr} -> {produced_ppr}",
                )
            )
        allowed = {
            json.dumps(run_signature(run), sort_keys=True, ensure_ascii=False)
            for run in expected.runs
        }
        for run in produced.runs:
            if is_drawing_run(run) or not run.text:
                continue
            signature = json.dumps(run_signature(run), sort_keys=True, ensure_ascii=False)
            if signature not in allowed:
                findings.append(
                    Finding(
                        rule["id"],
                        rule.get("severity", "error"),
                        "字符格式不在模板允许范围内",
                        location,
                        run.text[:40],
                        detail=f"{signature} not in {sorted(allowed)}",
                    )
                )
    return findings


def _check_table_shape(doc, template, rule) -> list[Finding]:
    if template is None:
        return []
    findings: list[Finding] = []
    if len(doc.tables) != len(template.tables):
        findings.append(
            Finding(
                rule["id"],
                rule.get("severity", "error"),
                "表格数量与模板不一致",
                {"kind": "table"},
                detail=f"template={len(template.tables)} output={len(doc.tables)}",
            )
        )
        return findings
    for index, (produced, expected) in enumerate(zip(doc.tables, template.tables)):
        if len(produced.columns) != len(expected.columns):
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "error"),
                    "表格列数与模板不一致",
                    {"kind": "table", "table": index},
                    detail=f"template={len(expected.columns)} output={len(produced.columns)}",
                )
            )
    return findings


def _check_image_fit(doc, template, rule) -> list[Finding]:
    findings: list[Finding] = []
    section = doc.sections[0] if doc.sections else None
    if section is None or section.page_width is None:
        return findings
    usable = int(section.page_width - (section.left_margin or 0) - (section.right_margin or 0))
    for index, shape in enumerate(doc.inline_shapes):
        width = int(shape.width)
        if width > usable:
            findings.append(
                Finding(
                    rule["id"],
                    rule.get("severity", "error"),
                    "图片宽度超过正文可用宽度",
                    {"kind": "body", "image": index},
                    detail=f"image={width} usable={usable} EMU",
                )
            )
    return findings




def _font_of(paragraph):
    if not paragraph.runs:
        return None
    font = paragraph.runs[0].font
    size = font.size.pt if font.size is not None else None
    return (font.name, size, bool(font.bold))


def _check_template_conformance(doc, template, rule) -> list[Finding]:
    """Compare the edited document with the original template layout."""
    if template is None:
        return []
    severity = rule.get("severity", "info")
    findings: list[Finding] = []
    template_paragraphs = list(template.paragraphs)
    output_paragraphs = list(doc.paragraphs)
    if len(template_paragraphs) != len(output_paragraphs):
        findings.append(
            Finding(
                rule["id"],
                severity,
                "段落数量与模板不同",
                {"kind": "body"},
                detail=f"template={len(template_paragraphs)} output={len(output_paragraphs)}",
            )
        )
    style_mismatch = []
    font_mismatch = []
    for index, (before, after) in enumerate(zip(template_paragraphs, output_paragraphs)):
        if (before.style.name or "") != (after.style.name or ""):
            style_mismatch.append((index, before.style.name, after.style.name))
        before_font = _font_of(before)
        after_font = _font_of(after)
        if before_font and after_font and before_font != after_font:
            font_mismatch.append((index, before_font, after_font))
    if style_mismatch:
        detail = "; ".join(f"p{index}: {a}→{b}" for index, a, b in style_mismatch[:6])
        findings.append(
            Finding(rule["id"], severity, "段落样式与模板不同", {"kind": "body"}, detail=detail)
        )
    if font_mismatch:
        detail = "; ".join(f"p{index}: {a}→{b}" for index, a, b in font_mismatch[:6])
        findings.append(
            Finding(rule["id"], severity, "字符格式与模板不同", {"kind": "body"}, detail=detail)
        )
    if len(doc.tables) != len(template.tables):
        findings.append(
            Finding(
                rule["id"],
                severity,
                "表格数量与模板不同",
                {"kind": "table"},
                detail=f"template={len(template.tables)} output={len(doc.tables)}",
            )
        )
    if template.sections and doc.sections:
        before = template.sections[0]
        after = doc.sections[0]
        diffs = [
            name
            for name in ("left_margin", "right_margin", "top_margin", "bottom_margin")
            if getattr(before, name) != getattr(after, name)
        ]
        if before.page_width != after.page_width or before.page_height != after.page_height:
            diffs.append("page_size")
        if diffs:
            findings.append(
                Finding(rule["id"], severity, "页面设置与模板不同", {"kind": "body"}, detail=",".join(diffs))
            )
    return findings


_HANDLERS = {
    "leftover_slot": _check_leftover_slot,
    "text_regex": _check_text_regex,
    "glued_words": _check_glued_words,
    "math_integrity": _check_math_integrity,
    "format_conformance": _check_conformance,
    "table_shape": _check_table_shape,
    "image_fit": _check_image_fit,
    "template_conformance": _check_template_conformance,
}


def run_checks(
    docx_path: str | Path,
    rules_path: str | Path | None = None,
    template_path: str | Path | None = None,
) -> list[Finding]:
    """Run every enabled rule; findings carry rule id, severity and location."""
    doc = Document(str(docx_path))
    template = Document(str(template_path)) if template_path else None
    # A template without markers is filled structurally, so extra paragraphs and
    # tables are expected; strict shape conformance would flag every such fill.
    structural = template is not None and not any(
        find_markers(paragraph) for paragraph in template.paragraphs
    )
    findings: list[Finding] = []
    for rule in load_rules(rules_path):
        if structural and rule["id"] in {"format-conformance", "table-shape"}:
            continue
        handler = _HANDLERS.get(rule["kind"])
        if handler is None:
            raise ValueError(f"unknown rule kind: {rule['kind']!r} in {rule['id']!r}")
        findings.extend(handler(doc, template, rule))
    return findings
