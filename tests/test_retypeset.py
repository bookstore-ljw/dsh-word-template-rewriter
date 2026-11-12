# -*- coding: utf-8 -*-
"""retypeset 的单元测试：用程序化生成的模板/资料，不依赖 Word 与真实论文。"""
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

from dsh_word.retypeset import (
    Profile,
    looks_placeholder,
    normalize_cross_run,
    retypeset,
)

TEMPLATE_PARAGRAPHS = [
    ("TTP Title", "Your Paper's Title Starts Here: Please Center"),
    ("TTP Author(s)", "John Keynes 1, and Adam Smith 2"),
    ("TTP Address", "1 School of Management, Some University, City, Country;"),
    ("TTP Abstract", "Abstract. Template abstract goes here."),
    ("TTP Keywords", "Keywords: alpha; beta; gamma."),
    ("Heading 1", "Introduction"),
    ("TTP Paragraph (others)", "Template body paragraph."),
    ("Heading 1", "References"),
]


def build_template(path: Path) -> None:
    document = Document()
    for style_name in ("figure caption", "equation", "Reference"):
        if style_name not in [style.name for style in document.styles]:
            document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    for style_name, text in TEMPLATE_PARAGRAPHS:
        if style_name not in [style.name for style in document.styles]:
            document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        document.add_paragraph(text, style=style_name)
    document.add_paragraph("Cheng Qiyun et al. A reference.", style="Reference")
    document.save(str(path))


def build_source(path: Path) -> None:
    document = Document()
    for style_name in ("paper title", "Author", "Abstract", "Keywords", "Heading 1"):
        if style_name not in [style.name for style in document.styles]:
            document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    for extra in ("figure caption", "table head", "equation", "Bibliography", "sponsors"):
        if extra not in [style.name for style in document.styles]:
            document.styles.add_style(extra, WD_STYLE_TYPE.PARAGRAPH)

    document.add_paragraph("A Real Title From Source", style="paper title")
    document.add_paragraph("Ada Lovelace, Charles Babbage", style="Author")
    document.add_paragraph("Abstract—This is the real abstract.", style="Abstract")
    document.add_paragraph("Keywords—one, two, three.", style="Keywords")
    document.add_paragraph("Introduction", style="Heading 1")
    document.add_paragraph("Body text before the table, see Table Ⅰ.", style="Body Text")
    document.add_paragraph("equation body paragraph", style="Body Text")
    equation = document.add_paragraph("", style="equation")
    equation.add_run("E = mc")
    document.add_paragraph("figure caption text", style="figure caption")
    document.add_paragraph("performance comparison", style="table head")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Metric"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Error"
    table.cell(1, 1).text = "0.1"
    document.add_paragraph("References", style="Heading 1")
    document.add_paragraph("[1]\tA. Author, “A paper,” 2024.", style="Bibliography")
    document.add_paragraph("[2]\tB. Author, “Another paper,” 2025.", style="Bibliography")
    document.save(str(path))


class ProfileTests(unittest.TestCase):
    def test_load_by_name(self) -> None:
        profile = Profile.load("ieee-to-ttp")
        self.assertEqual(profile.style("body"), "TTP Paragraph (others)")
        self.assertEqual(profile.heading_style(1), "Heading 1")
        self.assertEqual(profile.heading_style(2), "Heading 2")

    def test_unknown_profile_raises(self) -> None:
        with self.assertRaises(FileNotFoundError):
            Profile.load("no-such-profile")

    def test_placeholder_detection(self) -> None:
        self.assertTrue(looks_placeholder("Paper Title* (use style: paper title)"))
        self.assertTrue(looks_placeholder("1st Given Name Surname"))
        self.assertFalse(looks_placeholder("Physics-based tactile super-resolution"))


class CrossRunTests(unittest.TestCase):
    def test_table_reference_across_runs(self) -> None:
        document = Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("see Table ")
        paragraph.add_run("Ⅰ")
        normalize_cross_run(paragraph._p)
        self.assertEqual(paragraph.text, "see Table 1")


class RetypesetTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.mkdtemp()
        self.template = Path(self.tmp) / "template.docx"
        self.source = Path(self.tmp) / "source.docx"
        self.out = Path(self.tmp) / "out.docx"
        build_template(self.template)
        build_source(self.source)
        self.result = retypeset(self.template, self.source, self.out, "ieee-to-ttp")
        self.docx = Document(str(self.out))

    def test_counts(self) -> None:
        self.assertEqual(self.result.heading, 2)      # Introduction / References
        self.assertEqual(self.result.table, 1)
        self.assertEqual(self.result.reference, 2)
        self.assertGreaterEqual(self.result.paragraph, 2)

    def test_title_taken_from_source_when_real(self) -> None:
        texts = [p.text for p in self.docx.paragraphs]
        self.assertIn("A Real Title From Source", texts)
        self.assertNotIn("Your Paper's Title Starts Here: Please Center", texts)

    def test_abstract_and_keywords_lead(self) -> None:
        joined = "\n".join(p.text for p in self.docx.paragraphs)
        self.assertIn("Abstract. This is the real abstract.", joined)
        self.assertIn("Keywords: one, two, three.", joined)

    def test_reference_numbering_is_stripped(self) -> None:
        refs = [p.text for p in self.docx.paragraphs if p.style.name == "Reference"
                or p.style.name == "TTP Paragraph (others)" and p.text.startswith("A. Author")]
        joined = "\n".join(p.text for p in self.docx.paragraphs)
        self.assertIn("A. Author", joined)
        self.assertNotIn("[1]", joined)

    def test_captions_numbered(self) -> None:
        joined = "\n".join(p.text for p in self.docx.paragraphs)
        self.assertIn("Fig. 1 figure caption text", joined)
        self.assertIn("Table 1. Performance comparison", joined)

    def test_table_keeps_next_and_header(self) -> None:
        tbl = self.docx.tables[0]._tbl
        rows = tbl.findall(qn("w:tr"))
        self.assertIsNotNone(rows[0].find(qn("w:trPr")).find(qn("w:tblHeader")))
        self.assertIsNotNone(rows[0].find(qn("w:tc")).find(qn("w:p")).find(qn("w:pPr")).find(qn("w:keepNext")))

    def test_body_reference_normalized(self) -> None:
        joined = "\n".join(p.text for p in self.docx.paragraphs)
        self.assertIn("see Table 1", joined)

    def test_sponsors_dropped(self) -> None:
        joined = "\n".join(p.text for p in self.docx.paragraphs)
        self.assertNotIn("Identify applicable funding agency", joined)


if __name__ == "__main__":
    unittest.main()
