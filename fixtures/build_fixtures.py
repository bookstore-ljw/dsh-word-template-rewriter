"""Build a tiny but realistic fixture set: template, source, broken sample.

The template carries five slot kinds, each with a deliberate formatting trap:

* title        - Title style, 22pt, centered
* abstract     - 宋体 10.5pt, first-line indent expressed in characters
* method       - literal string slot inside a sentence
* image.fig1   - sized by the template (12cm), not by the source image
* formula.eq1  - display OMML copied from the source document
* table.refs   - row loop, two columns, template cell formatting inherited
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageDraw

MATH_XML = (
    "<m:oMathPara %s>"
    "<m:oMath>"
    "<m:r><m:t>E</m:t></m:r>"
    "<m:r><m:t>=</m:t></m:r>"
    "<m:r><m:t>m</m:t></m:r>"
    "<m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e>"
    "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>"
    "</m:oMath>"
    "</m:oMathPara>"
) % nsdecls("m", "w")


def _add_run(paragraph, text: str, size_pt: float = 12, bold: bool = False):
    run = paragraph.add_run(text)
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = "Times New Roman"
    rPr = run._element.get_or_add_rPr()
    fonts = rPr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = parse_xml("<w:rFonts %s/>" % nsdecls("w"))
        rPr.insert(0, fonts)
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    fonts.set(qn("w:eastAsia"), "宋体")
    return run


def _add_paragraph(doc, text: str = "", size_pt: float = 12, align=None, indent_chars=None):
    paragraph = doc.add_paragraph()
    if align is not None:
        paragraph.paragraph_format.alignment = align
    if indent_chars:
        ind = paragraph._p.get_or_add_pPr().get_or_add_ind()
        ind.set(qn("w:firstLineChars"), str(indent_chars * 100))
        ind.set(qn("w:firstLine"), str(int(indent_chars * size_pt * 20)))
    if text:
        _add_run(paragraph, text, size_pt)
    return paragraph


def build_template(path: Path) -> None:
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(2.54)
    section.left_margin = section.right_margin = Cm(3.18)

    title = doc.add_paragraph(style="Title")
    title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(title, "{{ text.title }}", 22, bold=True)

    _add_paragraph(doc, "摘要：{{ text.abstract }}", 10.5, indent_chars=2)

    doc.add_heading("1 方法", level=1)
    _add_paragraph(doc, "本文采用 {{ text.method }} 完成排版。", 12, indent_chars=2)

    caption = _add_paragraph(doc, "图 1 系统结构图", 10.5, WD_ALIGN_PARAGRAPH.CENTER)
    caption.runs[0].font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    _add_paragraph(doc, "{{ image.fig1 }}", 12, WD_ALIGN_PARAGRAPH.CENTER)
    _add_paragraph(doc, "{{ formula.eq1 }}", 12, WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_heading("2 参考文献", level=1)
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    header = ["编号", "文献条目"]
    for index, value in enumerate(header):
        _add_run(table.cell(0, index).paragraphs[0], value, 10.5, bold=True)
    marker_cell = table.cell(1, 0).paragraphs[0]
    _add_run(marker_cell, "{{ table.refs }}", 10.5)

    doc.save(str(path))


def build_source(path: Path, image: Path) -> None:
    doc = Document()
    _add_paragraph(doc, "基于模板填充的文档流水线", 16, WD_ALIGN_PARAGRAPH.CENTER)
    _add_paragraph(
        doc,
        "本文提出一种确定性的模板填充流水线，把资料内容复制到模板槽位，"
        "不重写任何句子，也不改动模板的段落与字符格式。",
        10.5,
    )
    _add_paragraph(doc, "填充层只做“取—放”两个动作，检查层只读不写。", 12)

    formula = doc.add_paragraph()
    formula._p.append(parse_xml(MATH_XML))

    doc.add_picture(str(image), width=Cm(8))

    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    rows = [
        ["[1]", "张三. 模板填充方法研究. 2024."],
        ["[2]", "Smith J. Deterministic template filling. 2023."],
    ]
    for row_index, values in enumerate(rows):
        for column_index, value in enumerate(values):
            table.cell(row_index, column_index).text = value
    doc.save(str(path))


def build_image(path: Path) -> None:
    image = Image.new("RGB", (800, 450), "#f4f7fb")
    draw = ImageDraw.Draw(image)
    draw.rectangle([40, 40, 760, 410], outline="#2f5d8a", width=4)
    draw.rectangle([90, 130, 250, 300], fill="#dbe7f3", outline="#2f5d8a", width=3)
    draw.rectangle([550, 130, 710, 300], fill="#f6e3c4", outline="#8a6a2f", width=3)
    draw.line([250, 215, 550, 215], fill="#2f5d8a", width=4)
    draw.text((120, 205), "source", fill="#1b1b1b")
    draw.text((585, 205), "template", fill="#1b1b1b")
    image.save(path, dpi=(300, 300))


def build_broken(path: Path) -> None:
    """Every failure mode the deterministic gates must catch, in one document."""
    doc = Document()
    _add_paragraph(doc, "The result is clear. Theresult is wrong.", 12)
    _add_paragraph(doc, "See figure 3.Next section for details.", 12)
    _add_paragraph(doc, "乱码字符\uFFFD出现在这里。", 12)
    _add_paragraph(doc, "残留占位符 {{ text.unknown }} 未替换。", 12)
    _add_paragraph(doc, "行内公式 $x^2 + y^2 = z^2$ 仍是 LaTeX。", 12)
    empty_math = doc.add_paragraph()
    empty_math._p.append(
        parse_xml("<m:oMathPara %s><m:oMath><m:r/></m:oMath></m:oMathPara>" % nsdecls("m", "w"))
    )
    doc.save(str(path))


PLAN = """\
version: 1
text:
  title: {doc: source.docx, paragraph: 0}
  abstract: {doc: source.docx, paragraph: 1}
  method: "模板槽位替换"
images:
  fig1: {path: fig1.png, width: 12cm}
formulas:
  eq1: {doc: source.docx, paragraph: 3}
tables:
  refs: {doc: source.docx, table: 0}
"""


def build(root: Path) -> dict[str, Path]:
    root.mkdir(parents=True, exist_ok=True)
    paths = {
        "root": root,
        "template": root / "template.docx",
        "source": root / "source.docx",
        "image": root / "fig1.png",
        "broken": root / "broken.docx",
        "plan": root / "plan.yaml",
    }
    build_template(paths["template"])
    build_image(paths["image"])
    build_source(paths["source"], paths["image"])
    build_broken(paths["broken"])
    # JSON-ish inline mappings in the YAML above are valid YAML flow mappings.
    paths["plan"].write_text(PLAN, encoding="utf-8")
    return paths


if __name__ == "__main__":
    result = build(Path(__file__).resolve().parent / "build")
    for key, value in result.items():
        print(f"{key}: {value}")
