"""Write manual-review entries for untouched UNO cells with ambiguous rich text."""
import csv
import os
import uuid
from pathlib import Path


def write_rich_text_review(output_path, entries):
    output = Path(output_path)
    report = output.with_name(output.stem + ".rich_text_review.csv")
    tmp = report.with_name(report.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with tmp.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(
                fh,
                fieldnames=["sheet", "object_id", "kind", "source_text", "reason"],
                extrasaction="ignore",
            )
            writer.writeheader()
            writer.writerows(entries)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, report)
    finally:
        if tmp.exists():
            tmp.unlink()
    return report
