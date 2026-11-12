"""Router tests: batching, patching one slot, and leaving the rest alone."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document

from dsh_word import fill_docx, load_plan, run_checks
from dsh_word.media import parse_length
from dsh_word.router import (
    Correction,
    RouterError,
    apply_batch,
    batch_corrections,
    load_corrections,
    run_router,
)
from fixtures.build_fixtures import build


class BatchTest(unittest.TestCase):
    def test_same_slot_merges_and_kinds_group(self) -> None:
        corrections = [
            Correction("image.fig1", {"width": "8cm"}, "偏小"),
            Correction("text.method", {"value": "A"}),
            Correction("image.fig1", {"width": "10cm"}, "再放大"),
            Correction("image.fig2", {"width": "6cm"}),
        ]
        batches = batch_corrections(corrections)
        self.assertEqual([batch.kind for batch in batches], ["image", "text"])
        image_batch = batches[0]
        self.assertEqual([c.slot for c in image_batch.corrections], ["image.fig1", "image.fig2"])
        merged = image_batch.corrections[0]
        self.assertEqual(merged.changes["width"], "10cm")
        self.assertIn("偏小", merged.note)
        self.assertIn("再放大", merged.note)

    def test_unknown_kind_is_rejected(self) -> None:
        with self.assertRaises(RouterError):
            batch_corrections([Correction("weird.thing", {"value": "x"})])

    def test_missing_text_value_is_rejected(self) -> None:
        with self.assertRaises(RouterError):
            batch_corrections([Correction("text.method", {})])


class ApplyTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.root = Path(tempfile.mkdtemp(prefix="dsh-word-router-"))
        cls.paths = build(cls.root / "fixtures")
        cls.plan = load_plan(cls.paths["plan"])

    def test_image_batch_resizes_without_touching_text(self) -> None:
        batch = batch_corrections(
            [Correction("image.fig1", {"width": "10cm"})]
        )[0]
        out = self.root / "image.docx"
        result = apply_batch(batch, self.paths["template"], self.plan, out)
        self.assertTrue(result.ok, result.as_dict())
        doc = Document(str(out))
        self.assertEqual(int(doc.inline_shapes[0].width), parse_length("10cm"))
        self.assertIn("模板槽位替换", doc.paragraphs[3].text)

    def test_text_batch_keeps_image_at_base_size(self) -> None:
        batch = batch_corrections(
            [Correction("text.method", {"value": "槽位替换（改）"})]
        )[0]
        out = self.root / "text.docx"
        result = apply_batch(batch, self.paths["template"], self.plan, out)
        self.assertTrue(result.ok, result.as_dict())
        doc = Document(str(out))
        self.assertIn("槽位替换（改）", doc.paragraphs[3].text)
        self.assertEqual(int(doc.inline_shapes[0].width), parse_length("12cm"))

    def test_router_writes_one_document_per_batch(self) -> None:
        corrections = self.root / "corrections.yaml"
        corrections.write_text(
            "corrections:\n"
            "  - {slot: image.fig1, width: 9cm}\n"
            "  - {slot: text.method, value: 合批测试}\n",
            encoding="utf-8",
        )
        summary = run_router(
            self.paths["template"],
            self.paths["plan"],
            corrections,
            self.root / "out",
            report_path=self.root / "report.json",
        )
        self.assertTrue(summary["ok"], summary)
        self.assertEqual(len(summary["results"]), 2)
        for result in summary["results"]:
            self.assertTrue(Path(result["output"]).exists())
            findings = run_checks(result["output"], template_path=self.paths["template"])
            self.assertEqual([f for f in findings if f.severity != 'info'], [])

    def test_dry_run_produces_no_files(self) -> None:
        corrections = self.root / "dry.yaml"
        corrections.write_text(
            "corrections:\n  - {slot: image.fig1, width: 9cm}\n", encoding="utf-8"
        )
        summary = run_router(
            self.paths["template"],
            self.paths["plan"],
            corrections,
            self.root / "dry-out",
            dry_run=True,
        )
        self.assertTrue(summary["dry_run"])
        self.assertEqual(summary["results"], [])
        self.assertFalse((self.root / "dry-out").exists())

    def test_load_corrections_keeps_notes(self) -> None:
        path = self.root / "notes.yaml"
        path.write_text(
            "corrections:\n  - {slot: image.fig1, width: 9cm, note: 按模板规定}\n",
            encoding="utf-8",
        )
        corrections = load_corrections(path)
        self.assertEqual(corrections[0].note, "按模板规定")
        self.assertEqual(corrections[0].changes["width"], "9cm")


if __name__ == "__main__":
    unittest.main(verbosity=2)
