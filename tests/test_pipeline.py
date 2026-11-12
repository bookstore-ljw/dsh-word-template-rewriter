"""End-to-end tests: copy fidelity, structural gates, tamper detection."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from dsh_word import FillError, FillPlan, fill_docx, load_plan, run_checks
from dsh_word.media import image_extent, parse_length
from dsh_word.omml import math_nodes, math_text
from dsh_word.values import ImageSpec, TableSpec
from fixtures.build_fixtures import build


class EndToEndTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.root = Path(tempfile.mkdtemp(prefix="dsh-word-"))
        cls.paths = build(cls.root / "fixtures")
        cls.plan = load_plan(cls.paths["plan"])
        cls.out = cls.root / "out.docx"
        cls.result = fill_docx(
            cls.paths["template"],
            cls.plan,
            cls.out,
            steps_dir=cls.root / "steps",
            trace_path=cls.root / "trace.json",
        )
        cls.doc = Document(str(cls.out))

    def test_plan_is_qualified_by_section(self) -> None:
        self.assertEqual(set(self.plan.text), {"text.title", "text.abstract", "text.method"})
        self.assertEqual(set(self.plan.images), {"image.fig1"})
        self.assertEqual(set(self.plan.formulas), {"formula.eq1"})
        self.assertEqual(set(self.plan.tables), {"table.refs"})

    def test_no_marker_survives(self) -> None:
        self.assertEqual(self.result.remaining, [])
        text = "\n".join(p.text for p in self.doc.paragraphs)
        self.assertNotIn("{{", text)
        self.assertNotIn("}}", text)

    def test_text_copied_verbatim(self) -> None:
        source = Document(str(self.paths["source"]))
        self.assertEqual(self.doc.paragraphs[0].text, source.paragraphs[0].text)
        self.assertIn("确定性的模板填充流水线", self.doc.paragraphs[1].text)
        self.assertIn("模板槽位替换", self.doc.paragraphs[3].text)

    def test_image_uses_template_size(self) -> None:
        self.assertEqual(len(self.doc.inline_shapes), 1)
        shape = self.doc.inline_shapes[0]
        self.assertEqual(int(shape.width), parse_length("12cm"))
        self.assertEqual(int(shape.height), parse_length("6.75cm"))

    def test_formula_is_omml_and_editable(self) -> None:
        nodes = [node for p in self.doc.paragraphs for node in math_nodes(p)]
        self.assertEqual(len(nodes), 1)
        self.assertEqual(math_text(nodes[0]), "E=mc2")
        self.assertIsNone(nodes[0].find(".//" + qn("w:t")))

    def test_table_rows_copied(self) -> None:
        table = self.doc.tables[0]
        self.assertEqual(len(table.rows), 3)
        self.assertEqual(table.cell(0, 0).text, "编号")
        self.assertEqual(table.cell(1, 0).text, "[1]")
        self.assertIn("Smith J.", table.cell(2, 1).text)

    def test_title_formatting_untouched(self) -> None:
        run = self.doc.paragraphs[0].runs[0]
        self.assertEqual(run.font.size, Pt(22))
        self.assertTrue(run.font.bold)
        fonts = run._element.rPr.rFonts
        self.assertEqual(fonts.get(qn("w:eastAsia")), "宋体")
        self.assertEqual(
            self.doc.paragraphs[0].paragraph_format.alignment, 1
        )  # WD_ALIGN_PARAGRAPH.CENTER

    def test_character_indent_preserved(self) -> None:
        ind = self.doc.paragraphs[1]._p.pPr.find(qn("w:ind"))
        self.assertEqual(ind.get(qn("w:firstLineChars")), "200")

    def test_steps_and_trace(self) -> None:
        self.assertEqual(len(self.result.steps), len(self.result.trace), "one step per op")
        self.assertEqual(len(self.result.steps), 6)
        for step in self.result.steps:
            self.assertTrue(Path(step).exists())
        self.assertEqual([item["kind"] for item in self.result.trace].count("text"), 3)
        self.assertTrue((self.root / "trace.json").exists())

    def test_clean_document_passes_gates(self) -> None:
        findings = run_checks(self.out, template_path=self.paths["template"])
        self.assertEqual([f.as_dict() for f in findings if f.severity != 'info'], [])


class CrossRunSlotTest(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(tempfile.mkdtemp(prefix="dsh-word-run-"))

    def _document_with_split_marker(self) -> Path:
        doc = Document()
        paragraph = doc.add_paragraph()
        bold = paragraph.add_run("前缀 {{ text.")
        bold.font.bold = True
        paragraph.add_run("title }} 后缀")
        path = self.root / "split.docx"
        doc.save(str(path))
        return path

    def test_marker_split_across_runs_keeps_format_and_suffix(self) -> None:
        template = self._document_with_split_marker()
        plan = FillPlan(text={"text.title": "已填充"})
        out = self.root / "split-out.docx"
        fill_docx(template, plan, out)
        paragraph = Document(str(out)).paragraphs[0]
        self.assertEqual(paragraph.text, "前缀 已填充 后缀")
        self.assertTrue(paragraph.runs[0].font.bold)

    def test_table_arity_mismatch_is_rejected(self) -> None:
        paths = build(self.root / "fixtures")
        plan = FillPlan(tables={"table.refs": TableSpec([["a", "b", "c"]])})
        with self.assertRaises(FillError):
            fill_docx(paths["template"], plan, self.root / "bad.docx")


class GateTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.root = Path(tempfile.mkdtemp(prefix="dsh-word-gate-"))
        cls.paths = build(cls.root / "fixtures")
        cls.plan = load_plan(cls.paths["plan"])
        cls.out = cls.root / "out.docx"
        fill_docx(cls.paths["template"], cls.plan, cls.out)

    def test_broken_document_hits_expected_rules(self) -> None:
        findings = run_checks(self.paths["broken"])
        rules = {finding.rule for finding in findings}
        self.assertTrue(
            {
                "no-leftover-slot",
                "replacement-character",
                "raw-tex",
                "math-integrity",
                "glued-words",
            }.issubset(rules),
            rules,
        )
        severity = {finding.rule: finding.severity for finding in findings}
        self.assertEqual(severity["replacement-character"], "error")
        self.assertEqual(severity["math-integrity"], "error")

    def test_tampered_font_is_caught(self) -> None:
        doc = Document(str(self.out))
        doc.paragraphs[0].runs[0].font.size = Pt(9)
        tampered = self.root / "tampered.docx"
        doc.save(str(tampered))
        findings = run_checks(tampered, template_path=self.paths["template"])
        self.assertIn("format-conformance", {finding.rule for finding in findings})

    def test_tampered_alignment_is_caught(self) -> None:
        doc = Document(str(self.out))
        doc.paragraphs[1].paragraph_format.alignment = 0
        tampered = self.root / "tampered-align.docx"
        doc.save(str(tampered))
        findings = run_checks(tampered, template_path=self.paths["template"])
        self.assertIn("format-conformance", {finding.rule for finding in findings})

    def test_oversized_image_is_caught(self) -> None:
        plan = load_plan(self.paths["plan"])
        width, height = image_extent(self.paths["image"], parse_length("20cm"))
        plan.images["image.fig1"] = ImageSpec(str(self.paths["image"]), width, height)
        out = self.root / "wide.docx"
        fill_docx(self.paths["template"], plan, out)
        findings = run_checks(out, template_path=self.paths["template"])
        self.assertIn("image-fit", {finding.rule for finding in findings})


if __name__ == "__main__":
    unittest.main(verbosity=2)
