import os
import shutil
import tempfile
import unittest

from app.core.filler import (ImageFillRequest, TextFillRequest,
                             CrossFillRequest, fill_from_images, fill_from_text,
                             fill_cross_files)
from app.core.models import ExtractionRule, FieldMapping, MappingEntry, TargetRef
from app.core.ocr import ManualOcrEngine


class FillerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import make_samples
        cls.samples = os.path.dirname(make_samples.__file__)
        # ensure samples exist
        for name in ("target.xlsx", "source.csv", "source.json", "template.docx",
                     "template.pptx", "table.png"):
            path = os.path.join(cls.samples, "samples", name)
            if not os.path.exists(path):
                make_samples.main()
        cls.tmp = tempfile.mkdtemp(prefix="tf_fill_")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def _out(self, name):
        return os.path.join(self.tmp, name)

    def test_text_fill_appends(self):
        target = os.path.join(self.samples, "samples", "target.xlsx")
        out = self._out("text_out.xlsx")
        shutil.copy2(target, out)
        report = fill_from_text(TextFillRequest(
            text="姓名：李四\n工号：E010\n部门：财务部\n入职日期：2022-07-15",
            target_path=out,
            output_path=out,
            append=True,
            column_plan=[
                FieldMapping(target="姓名", source="姓名"),
                FieldMapping(target="工号", source="工号"),
                FieldMapping(target="部门", source="部门"),
                FieldMapping(target="入职日期", source="入职日期"),
            ],
        ))
        self.assertEqual(report.written, 1)
        from app.core.spreadsheet import OpenpyxlEngine
        eng = OpenpyxlEngine(); eng.open(out)
        self.assertEqual(eng.cell_value("员工表", "A4"), "李四")
        self.assertEqual(eng.cell_value("员工表", "B4"), "E010")
        eng.close()

    def test_image_fields_mode(self):
        target = os.path.join(self.samples, "samples", "target.xlsx")
        out = self._out("img_fields.xlsx")
        shutil.copy2(target, out)
        words = [("李四", 10, 10, 40, 20), ("E010", 60, 10, 40, 20)]
        engine = ManualOcrEngine(words, text="姓名：李四\n工号：E010")
        report = fill_from_images(ImageFillRequest(
            images=[os.path.join(self.samples, "samples", "table.png")],
            target_path=out, output_path=out, mode="fields", append=True,
            column_plan=[FieldMapping(target="姓名", source="姓名"),
                         FieldMapping(target="工号", source="工号")],
        ), ocr_engine=engine)
        self.assertEqual(report.written, 1)

    def test_image_grid_mode(self):
        target = os.path.join(self.samples, "samples", "target.xlsx")
        out = self._out("img_grid.xlsx")
        shutil.copy2(target, out)
        # words positioned inside the detected 4x3 grid of table.png
        cells = [
            (0, 0, "姓名"), (1, 0, "工号"), (2, 0, "部门"), (3, 0, "入职日期"),
            (0, 1, "李四"), (1, 1, "E010"), (2, 1, "财务部"), (3, 1, "2022-07-15"),
        ]
        words = []
        for col, row, text in cells:
            x = col * 175 + 20
            y = row * 133 + 45
            words.append((text, x, y, 120, 40))
        engine = ManualOcrEngine(words, text="")
        report = fill_from_images(ImageFillRequest(
            images=[os.path.join(self.samples, "samples", "table.png")],
            target_path=out, output_path=out, mode="grid", append=True,
        ), ocr_engine=engine)
        self.assertGreater(report.written, 0)
        from app.core.spreadsheet import OpenpyxlEngine
        eng = OpenpyxlEngine(); eng.open(out)
        self.assertEqual(eng.cell_value("员工表", "A4"), "姓名")
        self.assertEqual(eng.cell_value("员工表", "B4"), "工号")
        self.assertEqual(eng.cell_value("员工表", "A5"), "李四")
        eng.close()

    def test_cross_files_xlsx(self):
        target = os.path.join(self.samples, "samples", "target.xlsx")
        out_dir = os.path.join(self.tmp, "cross_xlsx")
        os.makedirs(out_dir, exist_ok=True)
        report = fill_cross_files(CrossFillRequest(
            source_paths=[os.path.join(self.samples, "samples", "source.csv")],
            target_paths=[target],
            output_dir=out_dir,
            mappings=[
                MappingEntry(source="姓名", target=TargetRef(kind="xlsx_cell", cell="K1")),
                MappingEntry(source="工号", target=TargetRef(kind="xlsx_cell", cell="L1")),
            ],
        ))
        self.assertEqual(report.written, 6)  # 2 mappings x 3 source rows
        out_path = os.path.join(out_dir, "target.xlsx")
        self.assertTrue(os.path.exists(out_path))
        from app.core.spreadsheet import OpenpyxlEngine
        eng = OpenpyxlEngine(); eng.open(out_path)
        self.assertEqual(eng.cell_value("员工表", "K1"), "李四")
        self.assertEqual(eng.cell_value("员工表", "K2"), "王五")
        self.assertEqual(eng.cell_value("员工表", "K3"), "赵六")
        eng.close()

    def test_cross_files_docx_placeholder(self):
        target = os.path.join(self.samples, "samples", "template.docx")
        out_dir = os.path.join(self.tmp, "cross_docx")
        os.makedirs(out_dir, exist_ok=True)
        report = fill_cross_files(CrossFillRequest(
            source_paths=[os.path.join(self.samples, "samples", "source.csv")],
            target_paths=[target],
            output_dir=out_dir,
            batch_by_row=True,
            mappings=[
                MappingEntry(source="姓名", target=TargetRef(kind="docx_placeholder", placeholder="姓名")),
                MappingEntry(source="工号", target=TargetRef(kind="docx_placeholder", placeholder="工号")),
            ],
        ))
        self.assertEqual(report.written, 6)  # 2 mappings x 3 source rows
        first = os.path.join(out_dir, "template_001.docx")
        self.assertTrue(os.path.exists(first))
        import docx
        text = "\n".join(p.text for p in docx.Document(first).paragraphs)
        self.assertIn("李四", text)
        self.assertNotIn("{{姓名}}", text)

    def test_cross_files_pptx_table(self):
        target = os.path.join(self.samples, "samples", "template.pptx")
        out_dir = os.path.join(self.tmp, "cross_pptx")
        os.makedirs(out_dir, exist_ok=True)
        report = fill_cross_files(CrossFillRequest(
            source_paths=[os.path.join(self.samples, "samples", "source.csv")],
            target_paths=[target],
            output_dir=out_dir,
            batch_by_row=False,
            mappings=[
                MappingEntry(source="姓名", target=TargetRef(kind="pptx_table_cell", slide=0, table=0, row=1, col=0)),
                MappingEntry(source="工号", target=TargetRef(kind="pptx_table_cell", slide=0, table=0, row=1, col=1)),
            ],
        ))
        self.assertEqual(report.written, 2)
        from pptx import Presentation
        prs = Presentation(os.path.join(out_dir, "template.pptx"))
        table_shape = [s for s in prs.slides[0].shapes
                       if getattr(s, "has_table", False) and s.has_table][0]
        table = table_shape.table
        self.assertEqual(table.cell(1, 0).text, "李四")
        self.assertEqual(table.cell(1, 1).text, "E010")


if __name__ == "__main__":
    unittest.main()
