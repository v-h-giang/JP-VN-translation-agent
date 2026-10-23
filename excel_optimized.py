"""Sparse, cached and resumable Microsoft Excel COM translation pipeline."""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
from pathlib import Path

from excel_translate import (
    _load_excel_com, _iter_text_cells, _extract_cell_runs, _extract_shape_runs,
    _write_cell_runs, _write_shape_runs, _fit_shape_text, _safe_get,
)
from translation_batch import translatable, translate_context_batch

CHECKPOINT_INTERVAL = 75
CHECKPOINT_VERSION = 1
MSO_GROUP = 6


def digest(path):
    h = hashlib.sha256()
    with open(path, "rb") as stream:
        for part in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(part)
    return h.hexdigest()


def checkpoint_paths(output):
    return (
        output.with_name(output.stem + ".checkpoint" + output.suffix),
        output.with_name(output.stem + ".checkpoint.json"),
    )


def load_checkpoint(source_hash, output):
    book, manifest = checkpoint_paths(output)
    if not (book.exists() and manifest.exists()):
        return None
    try:
        state = json.loads(manifest.read_text(encoding="utf-8"))
        valid = (
            state.get("version") == CHECKPOINT_VERSION
            and state.get("source_hash") == source_hash
            and state.get("checkpoint_hash") == digest(book)
        )
        if not valid:
            print("Excel checkpoint hash/version mismatch: starting over.")
            return None
        print(f"Excel resume: {len(state['done'])} completed items")
        return state
    except (OSError, KeyError, ValueError, TypeError) as exc:
        print(f"Excel checkpoint ignored: {exc}")
        return None


def save_checkpoint(workbook, output, source_hash, done, cache, stats):
    book, manifest = checkpoint_paths(output)
    candidate = book.with_name(book.stem + ".writing" + book.suffix)
    manifest_tmp = manifest.with_name(manifest.name + ".writing")

    # SaveCopyAs leaves the active workbook, extension and formulas intact.
    workbook.SaveCopyAs(str(candidate))
    state = {
        "version": CHECKPOINT_VERSION,
        "source_hash": source_hash,
        "checkpoint_hash": digest(candidate),
        "done": sorted(done),
        "cache": cache,
        "stats": stats,
    }
    manifest_tmp.write_text(
        json.dumps(state, ensure_ascii=False), encoding="utf-8"
    )
    # Order matters: the new manifest references exactly the new copy's hash.
    os.replace(candidate, book)
    os.replace(manifest_tmp, manifest)
    print(f"Excel checkpoint: {len(done)} completed, {len(cache)} cached")


def iter_shapes(shapes, sheet_index, prefix=()):
    for i in range(1, int(shapes.Count) + 1):
        shape = shapes.Item(i)
        address = prefix + (i,)
        if int(_safe_get(shape, "Type", -1)) == MSO_GROUP:
            yield from iter_shapes(shape.GroupItems, sheet_index, address)
            continue
        try:
            frame = shape.TextFrame
            text = frame.Characters().Text
            if isinstance(text, str) and text.strip():
                yield {
                    "id": "S/{}/{}".format(
                        sheet_index, ".".join(map(str, address))
                    ),
                    "kind": "shape",
                    "target": shape,
                    "text": text,
                }
        except Exception:
            # Shapes without text (images/charts/OLE) are not translated.
            continue


def iter_items(workbook):
    for si in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets.Item(si)
        print(f"Excel sheet: {sheet.Name}")
        for cell in _iter_text_cells(sheet):
            address = str(cell.Address)
            yield {
                "id": f"C/{si}/{address}",
                "kind": "cell",
                "target": cell,
                "text": str(cell.Value2),
            }
        yield from iter_shapes(sheet.Shapes, si)


def runs_for(item):
    if item["kind"] == "cell":
        return _extract_cell_runs(item["target"], item["text"])
    return _extract_shape_runs(item["target"], item["text"])


def write_item(item, runs, stats):
    if item["kind"] == "cell":
        _write_cell_runs(item["target"], runs)
        stats["cells"] += 1
    else:
        _write_shape_runs(item["target"], runs)
        shrunk, resized = _fit_shape_text(item["target"], runs)
        stats["objects"] += 1
        stats["shrunk"] += int(shrunk)
        stats["resized"] += int(resized)


def flush(batch, cache, done, stats, translate_func):
    if not batch:
        return
    unique = []
    seen = set()
    for item in batch:
        for run in item["runs"]:
            text = run["text"]
            if not translatable(text):
                cache[text] = text
            elif text not in cache and text not in seen:
                seen.add(text)
                unique.append(text)

    if unique:
        translated = translate_context_batch(unique, translate_func)
        if len(translated) != len(unique):
            raise RuntimeError("COM batch produced wrong number of results")
        cache.update(zip(unique, translated))

    for item in batch:
        translated_runs = [
            {"text": cache[run["text"]], "style": run["style"]}
            for run in item["runs"]
        ]
        try:
            write_item(item, translated_runs, stats)
        except Exception as exc:
            raise RuntimeError(
                f"Excel COM failed writing {item['kind']} "
                f"{item['id']}: {type(exc).__name__}: {exc}"
            ) from exc
        done.add(item["id"])


def translate_excel_optimized(input_path, output_path, translate_func, max_chars=1200):
    """COM equivalent of UNO sparse scan, dedup, contextual batches and resume."""
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()
    if input_path == output_path:
        raise ValueError("Source and output must be different")
    if not input_path.is_file():
        raise FileNotFoundError(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    source_hash = digest(input_path)
    state = load_checkpoint(source_hash, output_path)
    book_checkpoint, manifest = checkpoint_paths(output_path)
    done = set(state["done"]) if state else set()
    cache = dict(state["cache"]) if state else {}
    stats = dict(state["stats"]) if state else {
        "cells": 0, "objects": 0, "shrunk": 0, "resized": 0
    }

    pythoncom, win32 = _load_excel_com()
    app = workbook = None
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix="jpvn_excel_") as temp:
            # Work only on staged bytes; keep source untouched.
            staging = Path(temp) / ("working" + input_path.suffix)
            shutil.copyfile(book_checkpoint if state else input_path, staging)

            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            app.EnableEvents = False
            app.ScreenUpdating = False
            try:
                app.AutomationSecurity = 3
            except Exception:
                pass
            workbook = app.Workbooks.Open(
                str(staging), UpdateLinks=0, ReadOnly=False,
                IgnoreReadOnlyRecommended=True, AddToMru=False,
            )
            batch = []
            pending_since_save = 0
            resumed = 0

            def save_progress():
                nonlocal batch, pending_since_save
                if not batch:
                    return
                flush(batch, cache, done, stats, translate_func)
                pending_since_save += len(batch)
                batch = []
                if pending_since_save >= CHECKPOINT_INTERVAL:
                    save_checkpoint(
                        workbook, output_path, source_hash,
                        done, cache, stats,
                    )
                    pending_since_save = 0

            for item in iter_items(workbook):
                if item["id"] in done:
                    resumed += 1
                    continue
                item["runs"] = runs_for(item)
                if not item["runs"]:
                    continue
                batch.append(item)
                if len(batch) >= CHECKPOINT_INTERVAL:
                    save_progress()
            save_progress()

            if pending_since_save:
                save_checkpoint(
                    workbook, output_path, source_hash,
                    done, cache, stats,
                )

            # Publish only a full workbook, never partial state.
            candidate = output_path.with_name(
                output_path.stem + ".complete" + output_path.suffix
            )
            workbook.SaveCopyAs(str(candidate))
            os.replace(candidate, output_path)
            print(
                f"Excel complete: {stats}; cache={len(cache)}; "
                f"resumed_skipped={resumed}"
            )
            for path in (book_checkpoint, manifest):
                try:
                    path.unlink()
                except FileNotFoundError:
                    pass
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
