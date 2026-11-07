"""End-to-end tests that exercise the real OCR + fill pipeline.

These are skipped automatically when RapidOCR is not installed, so the suite
still passes on a machine without OCR.
"""

import os
import shutil
import tempfile
import unittest

from app.core.filler import ImageFillRequest, fill_from_images
from app.core.models import FieldMapping


def rapidocr_available() -> bool:
    try:
        import rapidocr_onnxruntime  # noqa: F401
        return True
    except Exception:
        return False


@unittest.skipUnless(rapidocr_available(), "rapidocr-onnxruntime not installed")
class RapidOcrIntegrationTests(unittest.TestCase):
    def test_grid_fill_from_synthetic_table_image(self):
        import make_samples
        samples = os.path.dirname(make_samples.__file__)
        path = os.path.join(samples, "samples", "table.png")
        if not os.path.exists(path):
            make_samples.main()

        target = os.path.join(samples, "samples", "target.xlsx")
        tmp = tempfile.mkdtemp(prefix="tf_ocr_")
        out = os.path.join(tmp, "out.xlsx")
        shutil.copy2(target, out)

        report = fill_from_images(ImageFillRequest(
            images=[path],
            target_path=out,
            output_path=out,
            mode="grid",
            append=True,
        ))
        shutil.rmtree(tmp, ignore_errors=True)

        self.assertGreaterEqual(report.written, 3, msg="至少应识别出表格的 3 行")


if __name__ == "__main__":
    unittest.main()
