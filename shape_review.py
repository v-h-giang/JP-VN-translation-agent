"""Review report for shape font shrinking and resizing (both backends)."""
import csv
import os
import uuid
from pathlib import Path

HEADERS = [
    "backend", "sheet", "object_id", "object_name", "actions",
    "original_font_pt", "result_font_pt",
    "original_width", "result_width",
    "original_height", "result_height",
    "note",
]


def write_shape_review(output_path, entries):
    """Write UTF-8-with-BOM CSV only after final workbook is published."""
    output_path = Path(output_path)
    report = output_path.with_name(output_path.stem + ".shape_review.csv")
    pending = report.with_name(report.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with pending.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=HEADERS, extrasaction="ignore")
            writer.writeheader()
            for entry in entries:
                writer.writerow(entry)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(pending, report)
    finally:
        if pending.exists():
            pending.unlink()
    return report
