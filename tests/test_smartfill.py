import unittest
from pathlib import Path

from docx import Document

from dsh_word.smartfill import smart_fill


class SmartFillTests(unittest.TestCase):
    def test_smart_fill_keeps_template_structure(self) -> None:
        template = Path(self.tmp) / "template.docx"
        source = Path(self.tmp) / "source.docx"
        out = Path(self.tmp) / "out.docx"

        document = Document()
        document.add_heading("季度报告", level=1)
        document.add_paragraph("这里应该是摘要内容。")
        document.add_paragraph("这里应该是正文内容。")
        document.save(str(template))

        source_doc = Document()
        source_doc.add_paragraph("基于模板的文档流水线")
        source_doc.add_paragraph("摘要段落：只复制，不重写。")
        source_doc.add_paragraph("正文段落：检查层只读不写。")
        table = source_doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "项"
        table.cell(0, 1).text = "值"
        table.cell(1, 0).text = "步骤"
        table.cell(1, 1).text = "填充"
        source_doc.save(str(source))

        result = smart_fill(template, source, out)

        self.assertGreaterEqual(result.markers, 3)
        self.assertEqual(result.tables, 1)
        filled = Document(str(out))
        paragraphs = [paragraph.text for paragraph in filled.paragraphs]
        joined = "\n".join(paragraphs)
        # content comes from the source; the template heading and paragraph styles stay
        self.assertIn("基于模板的文档流水线", joined)
        self.assertIn("正文段落：检查层只读不写。", joined)
        self.assertIn("季度报告", joined)
        self.assertNotIn("这里应该是摘要内容。", joined)
        self.assertEqual(len(filled.tables), 1)
        self.assertEqual(filled.tables[0].cell(0, 1).text, "值")
        self.assertNotIn("{{", joined)

    def setUp(self) -> None:
        import tempfile

        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = self._tmp.name

    def tearDown(self) -> None:
        self._tmp.cleanup()


if __name__ == "__main__":
    unittest.main()
