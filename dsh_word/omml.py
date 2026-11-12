"""OMML helpers: extract source formulas, audit them, never re-typeset them."""

from __future__ import annotations

import copy
import re
from typing import Iterator

from docx.oxml.ns import qn

from .ooxml import M_OMATH, M_OMATHPARA, Location, iter_paragraphs

M_T = qn("m:t")
M_R = qn("m:r")
_VARIABLE_RE = re.compile(r"[A-Za-z][A-Za-z0-9]*")


def math_nodes(paragraph) -> list:
    """Direct OMML children of a paragraph (Word stores display math this way)."""
    return [child for child in paragraph._p if child.tag in (M_OMATH, M_OMATHPARA)]


def has_math(paragraph) -> bool:
    return bool(math_nodes(paragraph))


def extract_math(paragraph):
    """First math node in the paragraph, already deep-copied for insertion."""
    nodes = math_nodes(paragraph)
    if not nodes:
        raise ValueError("paragraph contains no OMML math")
    return copy.deepcopy(nodes[0])


def iter_math(doc) -> Iterator[tuple[Location, object]]:
    for location, paragraph in iter_paragraphs(doc):
        for node in math_nodes(paragraph):
            yield location, node


def math_text(node) -> str:
    return "".join(t.text or "" for t in node.iter(M_T))


def math_variables(node) -> list[str]:
    """Letter identifiers appearing inside math runs, for consistency checks."""
    text = math_text(node)
    return [token for token in _VARIABLE_RE.findall(text)]


def problems(node) -> list[str]:
    """Deterministic structural defects, phrased for a human report."""
    found: list[str] = []
    runs = list(node.iter(M_R))
    if not runs:
        found.append("公式没有 m:r 节点")
        return found
    texts = [t.text or "" for t in node.iter(M_T)]
    if not any(texts):
        found.append("公式文本为空")
    if len(texts) < len(runs):
        found.append(f"{len(runs)} 个 m:r 中只有 {len(texts)} 个带 m:t")
    joined = math_text(node)
    if "$" in joined:
        found.append("公式内残留 $ 定界符")
    if "\\" in joined:
        found.append("公式内残留 LaTeX 反斜杠")
    return found
