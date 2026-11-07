"""Generate sample workbooks / documents used by the tests and demos.

Run with::

    python make_samples.py
"""

from __future__ import annotations

import os

SAMPLES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples")


def ensure_dir() -> str:
    os.makedirs(SAMPLES, exist_ok=True)
    return SAMPLES


def make_target_xlsx() -> str:
    from openpyxl import Workbook
    path = os.path.join(SAMPLES, "target.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "员工表"
    ws.append(["姓名", "工号", "部门", "入职日期"])
    ws.append(["示例员工", "E001", "研发部", "2020-01-01"])
    ws.append(["张三", "E002", "市场部", "2021-03-05"])
    wb.create_sheet("空表")
    wb.save(path)
    return path


def make_source_csv() -> str:
    path = os.path.join(SAMPLES, "source.csv")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("姓名,工号,部门,入职日期\n")
        fh.write("李四,E010,财务部,2022-07-15\n")
        fh.write("王五,E011,研发部,2019-11-20\n")
        fh.write("赵六,E012,人事部,2023-02-01\n")
    return path


def make_source_json() -> str:
    import json
    path = os.path.join(SAMPLES, "source.json")
    data = [
        {"姓名": "钱七", "工号": "E020", "部门": "研发部", "入职日期": "2020-09-10"},
        {"姓名": "孙八", "工号": "E021", "部门": "市场部", "入职日期": "2021-12-01"},
    ]
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    return path


def make_docx() -> str:
    path = os.path.join(SAMPLES, "template.docx")
    import docx
    from docx.shared import Pt
    document = docx.Document()
    document.add_heading("入职信息登记", level=1)
    document.add_paragraph("姓名：{{姓名}}    工号：{{工号}}")
    document.add_paragraph("部门：{{部门}}    入职日期：{{入职日期}}")
    table = document.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "字段"
    hdr[1].text = "值"
    row = table.add_row().cells
    row[0].text = "姓名"
    row[1].text = "{{姓名}}"
    document.save(path)
    return path


def make_pptx() -> str:
    path = os.path.join(SAMPLES, "template.pptx")
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "员工信息"
    left, top, width, height = Inches(1), Inches(2), Inches(6), Inches(3)
    shape = slide.shapes.add_table(2, 2, left, top, width, height)
    table = shape.table
    table.cell(0, 0).text = "姓名"
    table.cell(0, 1).text = "工号"
    table.cell(1, 0).text = "{{姓名}}"
    table.cell(1, 1).text = "{{工号}}"
    prs.save(path)
    return path


def make_table_image() -> str:
    """Draw a small ruled table and save it as a PNG for OCR tests."""
    ensure_dir()
    path = os.path.join(SAMPLES, "table.png")
    import cv2
    import numpy as np

    width, height = 700, 400
    img = np.full((height, width, 3), 255, np.uint8)
    cols, rows = 4, 3
    cell_w, cell_h = width // cols, height // rows

    for c in range(cols + 1):
        cv2.line(img, (c * cell_w, 0), (c * cell_w, height), (0, 0, 0), 3)
    for r in range(rows + 1):
        cv2.line(img, (0, r * cell_h), (width, r * cell_h), (0, 0, 0), 3)

    labels = [
        ["姓名", "工号", "部门", "入职日期"],
        ["李四", "E010", "财务部", "2022-07-15"],
        ["王五", "E011", "研发部", "2019-11-20"],
    ]

    from PIL import Image, ImageDraw, ImageFont
    pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(pil)
    font_path = r"C:\Windows\Fonts\msyh.ttc"
    pil_font = None
    if os.path.exists(font_path):
        try:
            pil_font = ImageFont.truetype(font_path, 30)
        except Exception:
            pil_font = None
    for r, row in enumerate(labels):
        for c, text in enumerate(row):
            draw.text((c * cell_w + 14, r * cell_h + 8), text,
                      fill=(0, 0, 0), font=pil_font)
    pil.save(path)
    return path


def main() -> int:
    ensure_dir()
    paths = [
        make_target_xlsx(),
        make_source_csv(),
        make_source_json(),
        make_docx(),
        make_pptx(),
        make_table_image(),
    ]
    for path in paths:
        print("已生成", path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
