"""Automatic slot mapping: order-based, explicit about what it cannot map."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document

from dsh_word import fill_docx, run_checks
from dsh_word.autoplan import build_auto_plan
from fixtures.build_fixtures import build


class AutoPlanTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.root = Path(tempfile.mkdtemp(prefix="dsh-word-auto-"))
        cls.paths = build(cls.root / "fixtures")
        cls.auto = build_auto_plan(
            cls.paths["template"],
            cls.paths["source"],
            asset_dir=cls.root / "assets",
        )

    def test_every_template_slot_is_covered(self) -> None:
        self.assertEqual(self.auto.unresolved, [])
        slots = {item["slot"] for item in self.auto.mapping}
        self.assertEqual(
            slots,
            {
                "text.title",
                "text.abstract",
                "text.method",
                "image.fig1",
                "formula.eq1",
                "table.refs",
            },
        )

    def test_text_slots_follow_source_order(self) -> None:
        plan = self.auto.plan
        self.assertEqual(plan.text["text.title"], "基于模板填充的文档流水线")
        self.assertIn("确定性的模板填充流水线", plan.text["text.abstract"])
        self.assertIn("取—放", plan.text["text.method"])

    def test_image_is_extracted_to_disk(self) -> None:
        spec = self.auto.plan.images["image.fig1"]
        self.assertTrue(Path(spec.path).exists())
        self.assertGreater(spec.width, 0)

    def test_auto_plan_fills_and_passes_gates(self) -> None:
        out = self.root / "auto.docx"
        result = fill_docx(self.paths["template"], self.auto.plan, out)
        self.assertEqual(result.remaining, [])
        findings = run_checks(out, template_path=self.paths["template"])
        self.assertEqual([f for f in findings if f.severity != 'info'], [])
        doc = Document(str(out))
        self.assertEqual(len(doc.inline_shapes), 1)
        self.assertEqual(len(doc.tables[0].rows), 3)

    def test_missing_formula_is_reported(self) -> None:
        doc = Document()
        doc.add_paragraph("只有文字")
        stripped = self.root / "no-formula.docx"
        doc.save(str(stripped))
        auto = build_auto_plan(self.paths["template"], stripped, asset_dir=self.root / "a2")
        self.assertIn("formula.eq1", auto.unresolved)
        self.assertIn("image.fig1", auto.unresolved)


if __name__ == "__main__":
    unittest.main(verbosity=2)
