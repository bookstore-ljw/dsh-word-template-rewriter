# -*- coding: utf-8 -*-
"""按模板重排整篇文档（retypeset）。

与 ``fill`` / ``compose`` 的区别：

* ``fill``      —— 内容填进模板已有的段落槽位（模板结构不变）；
* ``compose``   —— 只借模板的样式名，重建标题/正文/公式/图/表，粒度较粗；
* ``retypeset`` —— **跨模板整篇重排**：模板只提供样式、页面设置、编号体系与
  题注格式，正文全部来自资料源，run 级格式（斜体、上下标、公式 OMML）原样搬运。

典型场景：一篇 IEEE 双栏会议稿要改成 Trans Tech / 单栏期刊格式。

样式的对应关系由 profile（JSON）描述，见 ``dsh_word/profiles/*.json``。
"""

from __future__ import annotations

import copy
import json
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm
from docx.text.paragraph import Paragraph

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

PROFILE_DIR = Path(__file__).with_name("profiles")

TABLE_REF_PATTERNS = [
    (re.compile("Table\\s*Ⅰ"), "Table 1"),
    (re.compile("Table\\s*Ⅱ"), "Table 2"),
    (re.compile("Table\\s+I\\b"), "Table 1"),
    (re.compile("Table\\s+II\\b"), "Table 2"),
]

# Word 对 pPr / rPr 等容器的子元素顺序敏感，最后统一按 schema 顺序重排
CHILD_ORDER = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
            "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
            "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
            "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
            "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
            "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
            "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
            "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl",
            "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
             "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins",
             "del", "trPrChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
             "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
             "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
              "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
              "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
              "tblPrChange"],
}

DROP_RPR = {"rFonts", "sz", "szCs", "kern", "spacing", "position", "w"}

PLACEHOLDER_HINTS = (
    "use style",
    "your paper",
    "paper title",
    "given name",
    "surname",
    "xxxxxxx",
    "template",
    "affiliation",
)


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def m(tag: str) -> str:
    return f"{{{M}}}{tag}"


def local(el) -> str:
    tag = el.tag if hasattr(el, "tag") else el
    return tag.rsplit("}", 1)[-1]


# --------------------------------------------------------------------------- #
# profile
# --------------------------------------------------------------------------- #
@dataclass
class Profile:
    target: dict
    source: dict
    layout: dict = field(default_factory=dict)
    flags: dict = field(default_factory=dict)

    @classmethod
    def load(cls, name_or_path: str | None) -> "Profile":
        raw = name_or_path or "ieee-to-ttp"
        candidate = Path(raw)
        if not candidate.exists():
            stem = raw if raw.endswith(".json") else f"{raw}.json"
            candidate = PROFILE_DIR / stem
        if not candidate.exists():
            raise FileNotFoundError(f"找不到 profile：{raw}（在 {PROFILE_DIR} 下也没有）")
        data = json.loads(candidate.read_text(encoding="utf-8"))
        return cls(
            target=data["target"],
            source=data["source"],
            layout=data.get("layout", {}),
            flags=data.get("flags", {}),
        )

    def style(self, key: str, default: str | None = None) -> str | None:
        value = self.target.get(key, default)
        if isinstance(value, dict):
            raise ValueError(f"{key} 是分级样式表，请用 heading_style()")
        return value

    def heading_style(self, level: int) -> str | None:
        table = self.target.get("heading", {})
        return table.get(str(level)) or table.get(level) or self.style("body")

    def opt(self, key: str, default):
        return self.flags.get(key, default)


def looks_placeholder(text: str) -> bool:
    lowered = (text or "").lower()
    return any(hint in lowered for hint in PLACEHOLDER_HINTS)


# --------------------------------------------------------------------------- #
# 低层工具
# --------------------------------------------------------------------------- #
def escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def make_run(text: str, rpr=None, red: bool = False):
    run = parse_xml(f'<w:r {nsdecls("w")}><w:t xml:space="preserve">{escape(text)}</w:t></w:r>')
    if rpr is not None:
        run.insert(0, copy.deepcopy(rpr))
    if red:
        set_color(run, "FF0000")
    return run


def make_tab_run():
    return parse_xml(f"<w:r {nsdecls('w')}><w:tab/></w:r>")


def get_or_make_rpr(run):
    rpr = run.find(q("rPr"))
    if rpr is None:
        rpr = run.makeelement(q("rPr"), {})
        run.insert(0, rpr)
    return rpr


def set_color(run, val: str) -> None:
    rpr = get_or_make_rpr(run)
    for old in rpr.findall(q("color")):
        rpr.remove(old)
    rpr.append(rpr.makeelement(q("color"), {q("val"): val}))


def redden_runs(p_el) -> None:
    for run in p_el.iter(q("r")):
        if run.find(q("t")) is not None or run.find(q("tab")) is not None:
            set_color(run, "FF0000")


def strip_ids(p_el) -> None:
    for attr in (f"{{{W14}}}paraId", f"{{{W14}}}textId"):
        if attr in p_el.attrib:
            del p_el.attrib[attr]


def proto_para(doc: Document, style_name: str | None):
    """取模板里第一个使用该样式的段落，只留 pPr 作为段落原型。"""
    if style_name is None:
        raise ValueError("原型样式名为空")
    for paragraph in doc.paragraphs:
        if paragraph.style is not None and paragraph.style.name == style_name:
            p_el = copy.deepcopy(paragraph._p)
            for child in list(p_el.iterchildren()):
                if child.tag != q("pPr"):
                    p_el.remove(child)
            strip_ids(p_el)
            return p_el
    raise KeyError(f"模板里没有样式为 {style_name!r} 的段落，无法取段落原型")


def proto_style_runs(doc: Document, style_name: str, max_runs: int = 2) -> list:
    for paragraph in doc.paragraphs:
        if paragraph.style is not None and paragraph.style.name == style_name:
            rprs = []
            for run in paragraph.runs[:max_runs]:
                rpr = run._r.find(q("rPr"))
                rprs.append(copy.deepcopy(rpr) if rpr is not None else None)
            return rprs
    return []


def clean_rpr(run) -> None:
    rpr = run.find(q("rPr"))
    if rpr is None:
        return
    for child in list(rpr.iterchildren()):
        name = local(child)
        if name in DROP_RPR:
            rpr.remove(child)
        elif name == "color" and (child.get(q("val")) or "").upper() != "FF0000":
            rpr.remove(child)
    if len(rpr) == 0:
        run.remove(rpr)


def compact_text(p) -> str:
    return re.sub(r"[ \t\r\n]+", " ", p.text).strip()


def has_image(p) -> bool:
    return p._p.find(".//" + q("drawing")) is not None


def blip_ids(p_el) -> list[str]:
    ids = []
    for blip in p_el.iter(f"{{{A}}}blip"):
        rid = blip.get(f"{{{R}}}embed")
        if rid:
            ids.append(rid)
    return ids


def transfer_content(src_p, tgt_p) -> None:
    """按顺序搬运 run / 公式；清掉源字体字号，让模板样式生效。"""
    for child in src_p.iterchildren():
        tag = local(child)
        if tag in ("pPr", "proofErr", "bookmarkStart", "bookmarkEnd",
                   "commentRangeStart", "commentRangeEnd", "sectPr"):
            continue
        if tag == "r":
            if child.find(q("t")) is None and child.find(q("tab")) is None and child.find(q("br")) is None:
                continue
            run = copy.deepcopy(child)
            for attr in list(run.attrib):
                if "rsid" in attr.lower():
                    del run.attrib[attr]
            clean_rpr(run)
            tgt_p.append(run)
        elif tag in ("oMath", "oMathPara"):
            tgt_p.append(copy.deepcopy(child))
        elif tag == "hyperlink":
            for sub in child.iterchildren():
                if local(sub) == "r":
                    run = copy.deepcopy(sub)
                    clean_rpr(run)
                    tgt_p.append(run)


def normalize_order(root) -> None:
    for tag, order in CHILD_ORDER.items():
        rank = {q(name): i for i, name in enumerate(order)}
        limit = len(order)
        for element in root.iter(q(tag)):
            children = list(element)
            ordered = sorted(children, key=lambda c: rank.get(c.tag, limit))
            if [id(c) for c in ordered] != [id(c) for c in children]:
                for c in children:
                    element.remove(c)
                for c in ordered:
                    element.append(c)


def normalize_cross_run(node, patterns=TABLE_REF_PATTERNS) -> None:
    """跨 run 的文本规范化（Word 常把 'Table ' 和罗马数字拆成两个 run）。"""
    nodes = [t for t in node.iter(q("t")) if t.text]
    if not nodes:
        return
    full = "".join(t.text for t in nodes)
    new = full
    for pattern, repl in patterns:
        new = pattern.sub(repl, new)
    if new == full:
        return
    cursor = 0
    for i, t in enumerate(nodes):
        if i == len(nodes) - 1:
            t.text = new[cursor:]
        else:
            t.text = new[cursor:cursor + len(t.text)]
            cursor += len(t.text)


def set_spacing(p_el, before=None, after=None, keep_next=False) -> None:
    ppr = p_el.find(q("pPr"))
    if ppr is None:
        ppr = p_el.makeelement(q("pPr"), {})
        p_el.insert(0, ppr)
    if before is not None or after is not None:
        for old in ppr.findall(q("spacing")):
            ppr.remove(old)
        attrs = {}
        if before is not None:
            attrs[q("before")] = str(before)
        if after is not None:
            attrs[q("after")] = str(after)
        ppr.append(ppr.makeelement(q("spacing"), attrs))
    if keep_next and ppr.find(q("keepNext")) is None:
        ppr.append(ppr.makeelement(q("keepNext"), {}))


def set_center_noindent(p_el) -> None:
    ppr = p_el.find(q("pPr"))
    if ppr is None:
        ppr = p_el.makeelement(q("pPr"), {})
        p_el.insert(0, ppr)
    for tag in ("jc", "ind"):
        for old in ppr.findall(q(tag)):
            ppr.remove(old)
    ppr.append(ppr.makeelement(q("jc"), {q("val"): "center"}))
    ppr.append(ppr.makeelement(q("ind"), {q("firstLine"): "0"}))


# --------------------------------------------------------------------------- #
# 表格
# --------------------------------------------------------------------------- #
def fit_table_widths(tbl, target_twip: int, font_pt: int = 10, margin_twip: int = 216) -> None:
    """按单元格文字的真实字宽分配列宽，避免表头/long word 被折行。"""
    try:
        import pymupdf
    except ImportError:  # 没有 PyMuPDF 时退化为等比缩放
        _rescale_table(tbl, target_twip)
        return

    regular = pymupdf.Font("tiro")
    bold = pymupdf.Font("tibo")

    rows = tbl.findall(q("tr"))
    if not rows:
        return
    ncols = len(rows[0].findall(q("tc")))
    needs = [margin_twip] * ncols
    for ri, tr in enumerate(rows):
        font = bold if ri == 0 else regular
        for ci, tc in enumerate(tr.findall(q("tc"))):
            if ci >= ncols:
                continue
            text = "".join(t.text or "" for t in tc.iter(q("t"))).strip()
            if not text:
                continue
            width_twip = font.text_length(text, fontsize=font_pt) * 20 * 1.04
            needs[ci] = max(needs[ci], width_twip + margin_twip)

    total = sum(needs)
    if total < target_twip:
        extra = target_twip - total
        needs = [w + extra * w / total for w in needs]
        total = target_twip
    factor = 9638 / total if total > 9638 else 1.0   # 不超过 A4 版心
    widths = [int(round(w * factor)) for w in needs]

    grid = tbl.find(q("tblGrid"))
    cols = grid.findall(q("gridCol")) if grid is not None else []
    for c, w in zip(cols, widths):
        c.set(q("w"), str(w))
    for tr in rows:
        for tc, w in zip(tr.findall(q("tc")), widths):
            tcpr = tc.find(q("tcPr"))
            if tcpr is None:
                tcpr = tc.makeelement(q("tcPr"), {})
                tc.insert(0, tcpr)
            for old in tcpr.findall(q("tcW")):
                tcpr.remove(old)
            tcpr.insert(0, tcpr.makeelement(q("tcW"), {q("w"): str(w), q("type"): "dxa"}))
    tblpr = tbl.find(q("tblPr"))
    if tblpr is not None:
        for tw in tblpr.findall(q("tblW")):
            tw.set(q("type"), "dxa")
            tw.set(q("w"), str(sum(widths)))


def _rescale_table(tbl, target_twip: int) -> None:
    grid = tbl.find(q("tblGrid"))
    cols = grid.findall(q("gridCol")) if grid is not None else []
    widths = [int(c.get(q("w"))) for c in cols if c.get(q("w"))]
    total = sum(widths)
    if total <= 0:
        return
    factor = target_twip / total
    for c, w in zip(cols, widths):
        c.set(q("w"), str(int(round(w * factor))))
    for tcw in tbl.iter(q("tcW")):
        if tcw.get(q("type")) in (None, "dxa") and tcw.get(q("w")):
            tcw.set(q("w"), str(int(round(int(tcw.get(q("w"))) * factor))))
    tblpr = tbl.find(q("tblPr"))
    if tblpr is not None:
        for tw in tblpr.findall(q("tblW")):
            tw.set(q("type"), "dxa")
            tw.set(q("w"), str(target_twip))


def polish_table(tbl, keep_next: bool = True, font_half_points: int = 20) -> None:
    """行不跨页 + 首行表头重复 + （可选）与下段同页，字号统一。"""
    for sz in tbl.iter(q("sz")):
        sz.set(q("val"), str(font_half_points))
    for sz in tbl.iter(q("szCs")):
        sz.set(q("val"), str(font_half_points))

    rows = tbl.findall(q("tr"))
    for i, tr in enumerate(rows):
        trpr = tr.find(q("trPr"))
        if trpr is None:
            trpr = tr.makeelement(q("trPr"), {})
            tr.insert(0, trpr)
        if trpr.find(q("cantSplit")) is None:
            trpr.append(trpr.makeelement(q("cantSplit"), {}))
        if i == 0 and trpr.find(q("tblHeader")) is None:
            trpr.append(trpr.makeelement(q("tblHeader"), {}))
        if i < len(rows) - 1 and keep_next:
            for p in tr.iter(q("p")):
                ppr = p.find(q("pPr"))
                if ppr is None:
                    ppr = p.makeelement(q("pPr"), {})
                    p.insert(0, ppr)
                if ppr.find(q("keepNext")) is None:
                    ppr.append(ppr.makeelement(q("keepNext"), {}))


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
@dataclass
class RetypesetResult:
    out_path: str
    heading: int = 0
    paragraph: int = 0
    equation: int = 0
    figure: int = 0
    table: int = 0
    reference: int = 0
    notes: list[str] = field(default_factory=list)
    styles_missing: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "out": self.out_path,
            "heading": self.heading,
            "paragraph": self.paragraph,
            "equation": self.equation,
            "figure": self.figure,
            "table": self.table,
            "reference": self.reference,
            "notes": self.notes,
            "styles_missing": self.styles_missing,
        }


def _sentence_case(text: str) -> str:
    text = text.strip()
    if not text:
        return text
    if text.isupper():
        return text[0] + text[1:].lower()
    return text[0].upper() + text[1:]


def _style_name(paragraph) -> str:
    return paragraph.style.name if paragraph.style is not None else ""


def retypeset(
    template_path: str | Path,
    source_path: str | Path,
    out_path: str | Path,
    profile: Profile | str | None = None,
    on_event=None,
) -> RetypesetResult:
    if not isinstance(profile, Profile):
        profile = Profile.load(profile)

    template_path, source_path, out_path = Path(template_path), Path(source_path), Path(out_path)
    tpl = Document(str(template_path))
    src = Document(str(source_path))
    assets = out_path.parent / "assets"
    result = RetypesetResult(out_path=str(out_path))

    figure_width_cm = float(profile.layout.get("figure_width_cm", 12.0))
    table_width_twip = int(profile.layout.get("table_width_twip", 6740))
    table_font_pt = int(profile.layout.get("table_font_pt", 10))
    keep_next = bool(profile.opt("keep_next", True))
    red_header = bool(profile.opt("red_placeholder_header", True))
    number_captions = bool(profile.opt("number_captions", True))
    normalize_refs = bool(profile.opt("normalize_table_refs", True))
    emit = on_event or (lambda event: None)

    # ---- 模板原型 --------------------------------------------------------
    tgt = profile.target
    src_cfg = profile.source

    def proto_or(style_name, fallback=None, required: bool = False):
        try:
            return proto_para(tpl, style_name)
        except (KeyError, ValueError):
            if required:
                raise
            result.styles_missing.append(str(style_name))
            return fallback

    P_BODY = proto_or(tgt.get("body"), required=True)
    P_TITLE = proto_or(tgt.get("title"), P_BODY)
    P_CAPTION = proto_or(tgt.get("caption") or tgt.get("body"), P_BODY)
    P_REF = proto_or(tgt.get("reference"), P_BODY)
    P_EQUATION = proto_or(tgt.get("equation"), P_BODY)
    P_ABSTRACT = proto_or(tgt.get("abstract"), None) if tgt.get("abstract") else None
    P_KEYWORDS = proto_or(tgt.get("keywords"), None) if tgt.get("keywords") else None
    P_REF_HEADING = proto_or(tgt.get("reference_heading"), None) if tgt.get("reference_heading") else None
    if P_REF_HEADING is not None and profile.opt("reference_heading_unumbered", True):
        ppr = P_REF_HEADING.find(q("pPr"))
        if ppr is not None:
            for num in ppr.findall(q("numPr")):
                ppr.remove(num)

    P_HEADING: dict[int, object] = {}
    for level in (1, 2, 3, 4):
        style_name = profile.heading_style(level)
        proto = proto_or(style_name, None)
        if proto is not None:
            P_HEADING[level] = proto
    if not P_HEADING:
        P_HEADING[1] = P_BODY

    R_ABS = proto_style_runs(tpl, tgt.get("abstract")) if tgt.get("abstract") else []
    R_KW = proto_style_runs(tpl, tgt.get("keywords")) if tgt.get("keywords") else []
    R_REF = proto_style_runs(tpl, tgt.get("reference") or tgt.get("body"), max_runs=1) or [None]

    # 标题/作者占位块：整段沿用模板（含上标），只标红
    header_paras = []
    if tgt.get("title"):
        for paragraph in tpl.paragraphs:
            if _style_name(paragraph) == tgt["title"]:
                p_el = copy.deepcopy(paragraph._p)
                strip_ids(p_el)
                if red_header:
                    redden_runs(p_el)
                header_paras.append(p_el)
                break
    author_styles = [s for s in (tgt.get("author"), tgt.get("address")) if s]
    skip_patterns = [re.compile(p, re.IGNORECASE)
                     for p in src_cfg.get("header_skip_patterns", [])]

    def is_instruction(paragraph) -> bool:
        text = paragraph.text or ""
        return any(pattern.search(text) for pattern in skip_patterns)

    in_author_block = False
    for paragraph in tpl.paragraphs:
        style = _style_name(paragraph)
        if style == tgt.get("author"):
            in_author_block = True
        if in_author_block and style in author_styles:
            if is_instruction(paragraph):
                result.notes.append(f"跳过模板作者块里的说明行：{compact_text(paragraph)[:30]}")
                continue
            p_el = copy.deepcopy(paragraph._p)
            strip_ids(p_el)
            if red_header:
                redden_runs(p_el)
            header_paras.append(p_el)
        elif in_author_block and style not in author_styles:
            in_author_block = False

    # 公式段落：居中公式 + 右对齐编号
    eq_ppr = P_EQUATION.find(q("pPr"))
    if eq_ppr is not None:
        for tag in ("jc", "tabs", "ind", "wordWrap"):
            for old in eq_ppr.findall(q(tag)):
                eq_ppr.remove(old)
        eq_ppr.append(eq_ppr.makeelement(q("wordWrap"), {q("val"): "0"}))
        eq_ppr.append(eq_ppr.makeelement(q("ind"), {q("firstLine"): "0"}))
        eq_ppr.append(eq_ppr.makeelement(q("jc"), {q("val"): "left"}))
        tabs = eq_ppr.makeelement(q("tabs"), {})
        tabs.append(tabs.makeelement(q("tab"), {q("val"): "center", q("pos"): "4820"}))
        tabs.append(tabs.makeelement(q("tab"), {q("val"): "right", q("pos"): "9640"}))
        eq_ppr.append(tabs)

    body = tpl.element.body
    for child in list(body.iterchildren()):
        if child.tag != q("sectPr"):
            body.remove(child)

    def add_para(proto):
        p_el = copy.deepcopy(proto)
        body.append(p_el)
        return p_el

    def add_text_run(p_el, text, rpr=None, red=False, bold=False):
        run = make_run(text, rpr, red=red)
        if bold:
            rpr_el = get_or_make_rpr(run)
            for name in ("b", "bCs"):
                if rpr_el.find(q(name)) is None:
                    rpr_el.append(rpr_el.makeelement(q(name), {}))
        p_el.append(run)
        return run

    src_cfg = profile.source
    title_styles = {s.lower() for s in src_cfg.get("title_styles", [])}
    abstract_styles = {s.lower() for s in src_cfg.get("abstract_styles", [])}
    keyword_styles = {s.lower() for s in src_cfg.get("keywords_styles", [])}
    fig_caption_styles = {s.lower() for s in src_cfg.get("figure_caption_styles", [])}
    tbl_caption_styles = {s.lower() for s in src_cfg.get("table_caption_styles", [])}
    equation_styles = {s.lower() for s in src_cfg.get("equation_styles", [])}
    reference_styles = {s.lower() for s in src_cfg.get("reference_styles", [])}
    drop_styles = {s.lower() for s in src_cfg.get("drop_styles", [])}
    drop_sections = [s.lower() for s in src_cfg.get("drop_sections", [])]
    author_style = (src_cfg.get("author_style") or "").lower()
    heading_prefix = (src_cfg.get("heading_prefix") or "Heading").lower()

    fig_no = eq_no = tbl_no = ref_no = 0
    saw_title = saw_author = False
    skip_section = False

    def save_image(rid, index):
        part = src.part.related_parts.get(rid)
        if part is None:
            return None
        assets.mkdir(parents=True, exist_ok=True)
        suffix = Path(str(part.partname)).suffix or ".png"
        path = assets / f"figure{index}{suffix}"
        path.write_bytes(part.blob)
        return path

    for block in src.element.body.iterchildren():
        tag = local(block)

        # ---- 表格 ----
        if tag == "tbl":
            if skip_section:
                result.notes.append("跳过：被排除章节内的表格")
                continue
            tbl_no += 1
            new = copy.deepcopy(block)
            for attr in list(new.attrib):
                if "rsid" in attr.lower() or "paraId" in attr or "textId" in attr:
                    del new.attrib[attr]
            fit_table_widths(new, table_width_twip, font_pt=table_font_pt)
            polish_table(new, keep_next=keep_next, font_half_points=table_font_pt * 2)
            body.append(new)
            result.table += 1
            emit({"kind": "table", "index": tbl_no})
            continue

        if tag != "p":
            continue

        paragraph = Paragraph(block, src)
        style = _style_name(paragraph)
        style_lower = style.lower()
        text = compact_text(paragraph)

        # ---- 排除章节 ----
        if style_lower.startswith(heading_prefix):
            lowered = text.lower()
            if any(key in lowered for key in drop_sections):
                skip_section = True
                result.notes.append(f"跳过章节：{text}")
                continue

        # ---- 标题层级 ----
        if style_lower.startswith(heading_prefix) and text:
            if src_cfg.get("reference_heading_text") and text.lower() == str(
                src_cfg["reference_heading_text"]
            ).lower():
                skip_section = False
                proto = P_REF_HEADING if P_REF_HEADING is not None else P_HEADING[1]
                add_text_run(add_para(proto), text)
                result.heading += 1
                emit({"kind": "heading", "text": text})
                continue
            skip_section = False
            digits = [c for c in style if c.isdigit()]
            level = int(digits[0]) if digits else 1
            proto = P_HEADING.get(level)
            if proto is None:
                proto = P_HEADING[1]
            add_text_run(add_para(proto), text)
            result.heading += 1
            emit({"kind": "heading", "text": text})
            continue

        if skip_section:
            continue

        # ---- 标题 ----
        if style_lower in title_styles:
            if saw_title:
                continue
            saw_title = True
            if text and not looks_placeholder(text):
                add_text_run(add_para(P_TITLE), text)
                result.notes.append(f"标题取自资料：{text[:60]}")
            else:
                for p_el in header_paras[:1]:
                    body.append(copy.deepcopy(p_el))
                result.notes.append(
                    f"资料标题是占位符（{text[:40]}），保留模板占位标题并标红"
                )
            continue

        # ---- 作者块 ----
        if author_style and style_lower == author_style and text.startswith("line 1"):
            if saw_author:
                continue
            saw_author = True
            if looks_placeholder(text) or not text:
                for p_el in header_paras[1:]:
                    body.append(copy.deepcopy(p_el))
                result.notes.append("资料作者/单位是占位符，保留模板占位并标红")
            else:
                proto_cls = proto_para(tpl, tgt.get("address") or tgt.get("author"))
                for line in [l.strip() for l in text.split("\n") if l.strip()]:
                    add_text_run(add_para(proto_cls), line)
                result.notes.append("作者块取自资料")
            continue

        # ---- 摘要 ----
        if style_lower in abstract_styles:
            proto = P_ABSTRACT if P_ABSTRACT is not None else P_BODY
            p_el = add_para(proto)
            lead = src_cfg.get("abstract_lead", "Abstract. ")
            body_text = re.sub(r"^(?:Abstract|ABSTRACT)\s*[—–\-.]\s*", "", text)
            add_text_run(p_el, lead, (R_ABS[0] if R_ABS else None), bold=True)
            add_text_run(p_el, body_text, (R_ABS[1] if len(R_ABS) > 1 else None))
            result.paragraph += 1
            emit({"kind": "abstract"})
            continue

        # ---- 关键词 ----
        if style_lower in keyword_styles:
            proto = P_KEYWORDS if P_KEYWORDS is not None else P_BODY
            p_el = add_para(proto)
            lead = src_cfg.get("keywords_lead", "Keywords:")
            words = re.sub(r"^(?:Keywords|KEYWORDS)\s*[—–\-:]\s*", "", text)
            add_text_run(p_el, lead, (R_KW[0] if R_KW else None), bold=True)
            add_text_run(p_el, " " + words, (R_KW[1] if len(R_KW) > 1 else None))
            result.paragraph += 1
            emit({"kind": "keywords"})
            continue

        # ---- 图 ----
        if has_image(paragraph):
            fig_no += 1
            p_el = add_para(P_CAPTION)
            set_center_noindent(p_el)
            set_spacing(p_el, before=120, after=0, keep_next=keep_next)
            para = Paragraph(p_el, tpl)
            for rid in blip_ids(block):
                saved = save_image(rid, fig_no)
                if saved is None:
                    continue
                para.add_run().add_picture(str(saved), width=Cm(figure_width_cm))
            result.figure += 1
            emit({"kind": "figure", "index": fig_no})
            continue

        # ---- 图题 / 表题 ----
        if style_lower in fig_caption_styles or style_lower in tbl_caption_styles:
            is_figure = style_lower in fig_caption_styles
            if is_figure:
                # 图题在图之后编号；若资料里图题孤悬（前面没有图），按下一张图预编号，
                # 避免出现 "Fig. 0"。
                figure_number = fig_no if fig_no else fig_no + 1
                label = f"Fig. {figure_number} " if number_captions else ""
            else:
                label = f"Table {tbl_no + 1}. " if number_captions else ""
            p_el = add_para(P_CAPTION)
            set_center_noindent(p_el)
            if is_figure:
                set_spacing(p_el, before=60, after=180)
            else:
                set_spacing(p_el, before=240, after=60, keep_next=keep_next)
            add_text_run(p_el, label + (_sentence_case(text) if not is_figure else text))
            result.notes.append(f"题注补编号：{label.strip()}")
            continue

        # ---- 独立公式 ----
        if style_lower in equation_styles:
            eq_no += 1
            p_el = add_para(copy.deepcopy(P_EQUATION))
            p_el.append(make_tab_run())
            for omath in block.iter(m("oMath")):
                p_el.append(copy.deepcopy(omath))
            p_el.append(make_tab_run())
            add_text_run(p_el, f"({eq_no})")
            result.equation += 1
            emit({"kind": "equation", "index": eq_no})
            continue

        # ---- 参考文献 ----
        if style_lower in reference_styles:
            if not text:
                continue
            ref_no += 1
            p_el = add_para(P_REF)
            add_text_run(p_el, re.sub(r"^\s*\[\d+\]\s*", "", text), R_REF[0])
            result.reference += 1
            emit({"kind": "reference", "index": ref_no})
            continue

        if style_lower in drop_styles:
            result.notes.append(f"跳过样式为 {style} 的段落（占位/资助声明等）")
            continue

        if not text:
            continue

        # ---- 正文 ----
        p_el = add_para(P_BODY)
        transfer_content(block, p_el)
        if normalize_refs:
            normalize_cross_run(p_el)
        result.paragraph += 1
        emit({"kind": "paragraph", "text": text[:60]})

    # 末尾空段清理
    while (len(body) >= 2 and local(body[-1]) == "p"
           and body[-1].find(".//" + q("drawing")) is None
           and not "".join(t.text or "" for t in body[-1].iter(q("t"))).strip()
           and local(body[-2]) == "p"):
        body.remove(body[-1])

    normalize_order(tpl.element)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tpl.save(str(out_path))
    return result


# --------------------------------------------------------------------------- #
# 版面真测（Word）
# --------------------------------------------------------------------------- #
def verify_layout(docx_path: str | Path, pdf_path: str | Path | None = None) -> dict:
    """用真实 Word 复核：页数、表格是否跨页、图与图题是否同页、PDF 是否带编辑标记。"""
    try:
        import win32com.client
    except ImportError as error:  # pragma: no cover
        raise RuntimeError("版面复核需要 Windows + pywin32") from error

    report: dict = {"page_count": 0, "tables": [], "figures": [], "issues": []}
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    try:
        document = word.Documents.Open(str(Path(docx_path).resolve()), False, True)
        document.Repaginate()
        report["page_count"] = int(document.ComputeStatistics(2))
        for i in range(1, document.Tables.Count + 1):
            table = document.Tables.Item(i)
            first = table.Range.Information(3)
            last = table.Rows.Item(table.Rows.Count).Range.Information(3)
            split = first != last
            report["tables"].append({"index": i, "rows": table.Rows.Count, "start_page": first,
                                     "end_page": last, "split": split})
            if split:
                report["issues"].append(f"表 {i} 跨页（{first}→{last}）")
        for i in range(1, document.InlineShapes.Count + 1):
            shape = document.InlineShapes.Item(i)
            page = shape.Range.Information(3)
            after = document.Range(shape.Range.End, document.Content.End).Paragraphs.Item(1)
            caption_page = after.Range.Information(3)
            ok = page == caption_page
            report["figures"].append({"index": i, "page": page, "caption_page": caption_page,
                                      "same_page": ok})
            if not ok:
                report["issues"].append(f"图 {i} 与图题不同页（{page}/{caption_page}）")
        if pdf_path is not None:
            document.ExportAsFixedFormat(str(Path(pdf_path).resolve()), 17)
        document.Close(0)
    finally:
        if document is not None:
            try:
                document.Close(0)
            except Exception:  # noqa: BLE001 - 已关闭时忽略
                pass
        word.Quit()

    if pdf_path is not None:
        try:
            import pymupdf
        except ImportError:
            pymupdf = None
        if pymupdf is not None:
            pdf = pymupdf.open(str(pdf_path))
            try:
                hits = sum(1 for page in pdf if "\u25a0" in page.get_text())
                report["pdf_pages"] = pdf.page_count
                report["pdf_editor_marks"] = hits
                if hits:
                    report["issues"].append(f"PDF 有 {hits} 页含编辑标记方块 ■")
            finally:
                pdf.close()
    return report
