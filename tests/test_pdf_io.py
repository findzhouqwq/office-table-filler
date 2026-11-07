import os
import shutil
import tempfile
import unittest

from app.core import pdf_io, data_io


def _make_pdf(path: str) -> None:
    import fitz
    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 72), "Name: Zhang San")
    page.insert_text((72, 100), "ID: E010")
    page2 = document.new_page()
    page2.insert_text((72, 72), "Dept: Finance")
    document.save(path)
    document.close()


@unittest.skipUnless(pdf_io.available(), "PyMuPDF not installed")
class PdfIoTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="tf_pdf_")
        self.pdf = os.path.join(self.tmp, "sample.pdf")
        _make_pdf(self.pdf)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_inspect_and_extract(self):
        info = pdf_io.inspect_pdf(self.pdf)
        self.assertEqual(info.pages, 2)
        self.assertTrue(info.has_text)
        text = pdf_io.extract_pdf_text(self.pdf)
        self.assertIn("Zhang San", text)
        self.assertIn("Finance", text)

    def test_render_pages(self):
        out_dir = os.path.join(self.tmp, "pages")
        pages = pdf_io.render_pdf_pages(self.pdf, out_dir=out_dir, dpi=72)
        self.assertEqual(len(pages), 2)
        for page in pages:
            self.assertTrue(os.path.isfile(page))
            self.assertGreater(os.path.getsize(page), 0)

    def test_expand_image_inputs(self):
        pages, temp_dir = pdf_io.expand_image_inputs(
            [self.pdf, os.path.join(self.tmp, "not-a-pdf.png")], max_pdf_pages=1)
        try:
            self.assertEqual(len(pages), 2)  # 1 PDF page + the plain path
            self.assertTrue(os.path.isdir(temp_dir))
        finally:
            pdf_io.cleanup_temp_dir(temp_dir)

    def test_load_table_reads_pdf(self):
        table = data_io.load_table(self.pdf)
        self.assertTrue(table.headers)
        joined = " ".join(str(cell) for row in table.rows for cell in row)
        self.assertIn("Zhang San", joined)
        self.assertIn("Finance", joined)


if __name__ == "__main__":
    unittest.main()
