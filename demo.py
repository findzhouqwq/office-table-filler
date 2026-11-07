"""Run all three fill features against the generated samples.

Usage::

    python make_samples.py
    python demo.py
"""

from __future__ import annotations

import os

from app.core.filler import (CrossFillRequest, ImageFillRequest, TextFillRequest,
                             fill_cross_files, fill_from_images, fill_from_text)
from app.core.models import FieldMapping, MappingEntry, TargetRef
from app.core.ocr import get_ocr_engine

ROOT = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(ROOT, "samples")
OUT = os.path.join(ROOT, "demo_output")
os.makedirs(OUT, exist_ok=True)


def line(title: str) -> None:
    print("\n" + "=" * 60)
    print(title)
    print("=" * 60)


def feature_1_images() -> None:
    line("① 图片识别批量填表（网格识别）")
    target = os.path.join(OUT, "feature1_target.xlsx")
    import shutil
    base = os.path.join(SAMPLES, "target.xlsx")
    shutil.copy2(base, target)

    image = os.path.join(SAMPLES, "table.png")
    if not os.path.exists(image):
        print("跳过：缺少样例图片")
        return
    engine = get_ocr_engine("rapidocr")
    report = fill_from_images(ImageFillRequest(
        images=[image],
        target_path=target,
        output_path=target,
        mode="grid",
        append=True,
    ), ocr_engine=engine)
    print(f"写入 {report.written} 处，引擎 {engine.name}")
    print("输出：", report.output_path)


def feature_2_text() -> None:
    line("② 文字填表（CSV → Excel）")
    target = os.path.join(OUT, "feature2_target.xlsx")
    import shutil
    shutil.copy2(os.path.join(SAMPLES, "target.xlsx"), target)
    source = os.path.join(SAMPLES, "source.csv")

    # Read the CSV and map columns by header name.
    from app.core.data_io import load_table
    table = load_table(source)
    plan = [FieldMapping(target=h, source=h) for h in table.headers]
    report = fill_from_text(TextFillRequest(
        records=table.as_records(),
        target_path=target,
        output_path=target,
        append=True,
        column_plan=plan,
    ))
    print(f"写入 {report.written} 处")
    print("输出：", report.output_path)


def feature_3_cross() -> None:
    line("③ 跨文件填表（源 CSV → Excel / Word / PPT）")
    source = os.path.join(SAMPLES, "source.csv")
    xlsx = os.path.join(SAMPLES, "template.xlsx") if os.path.exists(os.path.join(SAMPLES, "template.xlsx")) else os.path.join(SAMPLES, "target.xlsx")
    word = os.path.join(SAMPLES, "template.docx")
    power = os.path.join(SAMPLES, "template.pptx")
    targets = [xlsx]
    if os.path.exists(word):
        targets.append(word)
    if os.path.exists(power):
        targets.append(power)

    mappings = [
        MappingEntry(source="姓名", target=TargetRef(kind="xlsx_cell", cell="F1", row=1)),
        MappingEntry(source="工号", target=TargetRef(kind="xlsx_cell", cell="G1", row=1)),
    ]
    if os.path.exists(word):
        mappings.append(MappingEntry(source="姓名", target=TargetRef(kind="docx_placeholder", placeholder="姓名")))
    if os.path.exists(power):
        mappings.append(MappingEntry(source="姓名", target=TargetRef(kind="pptx_table_cell", slide=0, table=0, row=1, col=0)))
        mappings.append(MappingEntry(source="工号", target=TargetRef(kind="pptx_table_cell", slide=0, table=0, row=1, col=1)))

    report = fill_cross_files(CrossFillRequest(
        source_paths=[source],
        target_paths=targets,
        output_dir=OUT,
        batch_by_row=True,
        mappings=mappings,
    ))
    print(f"写入 {report.written} 处")
    for warning in report.warnings:
        print("  ·", warning)
    print("输出目录：", report.output_path)


def main() -> int:
    import make_samples
    make_samples.main()
    feature_1_images()
    feature_2_text()
    feature_3_cross()
    print("\n完成。结果在", OUT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
