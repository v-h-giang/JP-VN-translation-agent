from __future__ import annotations

import os
import re
import shutil
import stat
from pathlib import Path
from typing import Callable

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}

XL_CELL_TYPE_CONSTANTS = 2
XL_TEXT_VALUES = 2
MSO_GROUP = 6
MSO_TRUE = -1

MIN_SHAPE_FONT_SIZE = 8.0
FIT_TOLERANCE_PT = 1.0
MAX_FONT_FIT_STEPS = 6

PROTECTED_TOKEN_RE = re.compile(r"__PROTECTED_\d{6}__")


def _load_excel_com():
    if os.name != "nt":
        raise RuntimeError(
            "Excel translation requires Windows with Microsoft Excel installed."
        )

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "Excel translation requires pywin32. Run setup.bat again."
        ) from exc

    return pythoncom, win32com.client


def _safe_get(obj, name, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _safe_set(obj, name, value):
    if value is None:
        return

    try:
        setattr(obj, name, value)
    except Exception:
        pass


def _as_float(value):
    try:
        if value is None:
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _cell_font_style(font):
    return {
        "name": _safe_get(font, "Name"),
        "size": _as_float(_safe_get(font, "Size")),
        "bold": _safe_get(font, "Bold"),
        "italic": _safe_get(font, "Italic"),
        "underline": _safe_get(font, "Underline"),
        "color": _safe_get(font, "Color"),
        "strikethrough": _safe_get(font, "Strikethrough"),
        "superscript": _safe_get(font, "Superscript"),
        "subscript": _safe_get(font, "Subscript"),
    }


def _shape_font_style(font):
    # Excel Shape.TextFrame.Characters(...).Font returns the Excel Font
    # interface, which is more stable across Excel versions than Font2.
    return _cell_font_style(font)


def _style_looks_mixed(style):
    # These are the attributes the user explicitly cares about most.
    for key in ("name", "size", "bold", "italic", "color"):
        if style.get(key) is None:
            return True

    return False


def _merge_character_runs(text, style_at):
    runs = []
    current_style = None
    current_text = []

    for index, char in enumerate(text, start=1):
        style = style_at(index)

        if current_style is None:
            current_style = style
            current_text = [char]
            continue

        if style == current_style:
            current_text.append(char)
            continue

        runs.append(
            {
                "text": "".join(current_text),
                "style": current_style,
            }
        )
        current_style = style
        current_text = [char]

    if current_text:
        runs.append(
            {
                "text": "".join(current_text),
                "style": current_style,
            }
        )

    return runs


def _extract_cell_runs(cell, text):
    whole_style = _cell_font_style(cell.Font)

    if len(text) <= 1 or not _style_looks_mixed(whole_style):
        return [{"text": text, "style": whole_style}]

    return _merge_character_runs(
        text,
        lambda index: _cell_font_style(
            cell.Characters(index, 1).Font
        ),
    )


def _extract_shape_runs(shape, text):
    text_frame = shape.TextFrame
    whole_style = _shape_font_style(
        text_frame.Characters().Font
    )

    if len(text) <= 1 or not _style_looks_mixed(whole_style):
        return [{"text": text, "style": whole_style}]

    return _merge_character_runs(
        text,
        lambda index: _shape_font_style(
            text_frame.Characters(index, 1).Font
        ),
    )


def _apply_cell_font(font, style):
    _safe_set(font, "Name", style.get("name"))
    _safe_set(font, "Size", style.get("size"))
    _safe_set(font, "Bold", style.get("bold"))
    _safe_set(font, "Italic", style.get("italic"))
    _safe_set(font, "Underline", style.get("underline"))
    _safe_set(font, "Color", style.get("color"))
    _safe_set(font, "Strikethrough", style.get("strikethrough"))
    _safe_set(font, "Superscript", style.get("superscript"))
    _safe_set(font, "Subscript", style.get("subscript"))


def _apply_shape_font(font, style):
    _apply_cell_font(font, style)


def _split_text_for_limit(text, limit):
    if len(text) <= limit:
        return [text]

    units = re.split(r"(?<=[.!?。！？\n])", text)
    pieces = []
    buffer = ""

    for unit in units:
        if not unit:
            continue

        if len(buffer) + len(unit) <= limit:
            buffer += unit
            continue

        if buffer:
            pieces.append(buffer)
            buffer = ""

        if len(unit) <= limit:
            buffer = unit
            continue

        start = 0
        while start < len(unit):
            pieces.append(unit[start : start + limit])
            start += limit

    if buffer:
        pieces.append(buffer)

    return pieces


def _normalize_runs_for_limit(runs, max_chars):
    result = []
    text_limit = max(200, max_chars - 120)

    for run in runs:
        text = run["text"]

        if len(text) <= text_limit:
            result.append(run)
            continue

        for piece in _split_text_for_limit(text, text_limit):
            result.append(
                {
                    "text": piece,
                    "style": dict(run["style"]),
                }
            )

    return result


def _batch_runs(runs, max_chars):
    batches = []
    current = []
    current_size = 0

    for run in _normalize_runs_for_limit(runs, max_chars):
        estimated = len(run["text"]) + 50

        if current and current_size + estimated > max_chars:
            batches.append(current)
            current = []
            current_size = 0

        current.append(run)
        current_size += estimated

    if current:
        batches.append(current)

    return batches


def _encode_run_batch(runs):
    payload = []
    expected_tokens = []

    for index, run in enumerate(runs):
        token_base = 800000 + index * 2
        start_token = f"__PROTECTED_{token_base:06d}__"
        end_token = f"__PROTECTED_{token_base + 1:06d}__"

        expected_tokens.extend(
            [
                start_token,
                end_token,
            ]
        )

        payload.append(start_token)
        payload.append(run["text"])
        payload.append(end_token)

    return "".join(payload), expected_tokens


def _decode_run_batch(translated, runs, expected_tokens):
    actual_tokens = PROTECTED_TOKEN_RE.findall(translated)

    if actual_tokens != expected_tokens:
        raise RuntimeError(
            "TranslateGemma changed Excel rich-text run markers."
        )

    decoded = []
    cursor = 0

    for index, run in enumerate(runs):
        start_token = expected_tokens[index * 2]
        end_token = expected_tokens[index * 2 + 1]

        start_pos = translated.find(start_token, cursor)
        if start_pos < 0:
            raise RuntimeError("Missing Excel rich-text start marker.")

        if translated[cursor:start_pos]:
            raise RuntimeError(
                "Unexpected text outside Excel rich-text markers."
            )

        content_start = start_pos + len(start_token)
        end_pos = translated.find(end_token, content_start)

        if end_pos < 0:
            raise RuntimeError("Missing Excel rich-text end marker.")

        decoded.append(
            {
                "text": translated[content_start:end_pos],
                "style": dict(run["style"]),
            }
        )

        cursor = end_pos + len(end_token)

    if translated[cursor:]:
        raise RuntimeError(
            "Unexpected trailing text outside Excel rich-text markers."
        )

    return decoded


def translate_rich_runs(
    runs,
    translate_func: Callable[[str], str],
    max_chars,
):
    translated_runs = []

    for batch in _batch_runs(runs, max_chars):
        payload, expected_tokens = _encode_run_batch(batch)

        try:
            translated = translate_func(payload)
            decoded = _decode_run_batch(
                translated,
                batch,
                expected_tokens,
            )
            translated_runs.extend(decoded)
            continue

        except Exception as exc:
            print(
                "    rich-text marker fallback: "
                f"{type(exc).__name__}: {exc}"
            )

        # Rare fallback: translate each formatting run separately.
        # This preserves formatting exactly, at the cost of less sentence context.
        for run in batch:
            if not run["text"].strip():
                translated_text = run["text"]
            else:
                translated_text = translate_func(run["text"])

            translated_runs.append(
                {
                    "text": translated_text,
                    "style": dict(run["style"]),
                }
            )

    return translated_runs


def _write_cell_runs(cell, translated_runs):
    text = "".join(run["text"] for run in translated_runs)

    if len(text) > 32767:
        raise RuntimeError(
            "Translated cell text exceeds Excel's 32,767-character limit."
        )

    cell.Value2 = text

    start = 1

    for run in translated_runs:
        run_text = run["text"]

        if run_text:
            chars = cell.Characters(start, len(run_text))
            _apply_cell_font(
                chars.Font,
                run["style"],
            )

        start += len(run_text)


def _write_shape_runs(shape, translated_runs):
    text_frame = shape.TextFrame
    characters = text_frame.Characters()
    characters.Text = "".join(
        run["text"]
        for run in translated_runs
    )

    start = 1

    for run in translated_runs:
        run_text = run["text"]

        if run_text:
            font = text_frame.Characters(
                start,
                len(run_text),
            ).Font

            _apply_shape_font(
                font,
                run["style"],
            )

        start += len(run_text)


def _shape_metrics(shape):
    # TextRange2 BoundWidth/BoundHeight are documented Office members and
    # provide the rendered text bounds. Keep TextFrame2 usage isolated here.
    text_frame = shape.TextFrame2
    text_range = text_frame.TextRange

    margin_left = float(_safe_get(text_frame, "MarginLeft", 0) or 0)
    margin_right = float(_safe_get(text_frame, "MarginRight", 0) or 0)
    margin_top = float(_safe_get(text_frame, "MarginTop", 0) or 0)
    margin_bottom = float(_safe_get(text_frame, "MarginBottom", 0) or 0)

    width = float(shape.Width)
    height = float(shape.Height)

    available_width = max(
        1.0,
        width - margin_left - margin_right,
    )
    available_height = max(
        1.0,
        height - margin_top - margin_bottom,
    )

    bound_width = float(
        _safe_get(text_range, "BoundWidth", 0)
        or 0
    )
    bound_height = float(
        _safe_get(text_range, "BoundHeight", 0)
        or 0
    )

    return {
        "available_width": available_width,
        "available_height": available_height,
        "bound_width": bound_width,
        "bound_height": bound_height,
        "margin_width": margin_left + margin_right,
        "margin_height": margin_top + margin_bottom,
    }


def _shape_overflow(shape):
    metrics = _shape_metrics(shape)

    horizontal = (
        metrics["bound_width"]
        > metrics["available_width"] + FIT_TOLERANCE_PT
    )
    vertical = (
        metrics["bound_height"]
        > metrics["available_height"] + FIT_TOLERANCE_PT
    )

    return horizontal, vertical, metrics


def _apply_shape_font_scale(
    shape,
    translated_runs,
    scale,
):
    text_frame = shape.TextFrame
    start = 1

    for run in translated_runs:
        run_text = run["text"]
        original_size = run["style"].get("size")

        if run_text and original_size:
            target_size = max(
                MIN_SHAPE_FONT_SIZE,
                float(original_size) * scale,
            )

            text_frame.Characters(
                start,
                len(run_text),
            ).Font.Size = target_size

        start += len(run_text)


def _fit_shape_text(shape, translated_runs):
    sizes = [
        float(run["style"]["size"])
        for run in translated_runs
        if run["text"]
        and run["style"].get("size")
    ]

    shrunk = False
    resized = False

    horizontal, vertical, metrics = _shape_overflow(shape)

    if not horizontal and not vertical:
        return shrunk, resized

    # If the source already contains text <= 8pt, do not make it smaller.
    # Go directly to object resizing when necessary.
    can_shrink = bool(sizes) and min(sizes) > MIN_SHAPE_FONT_SIZE

    if can_shrink:
        min_scale = MIN_SHAPE_FONT_SIZE / min(sizes)
        current_scale = 1.0

        for _ in range(MAX_FONT_FIT_STEPS):
            horizontal, vertical, metrics = _shape_overflow(shape)

            if not horizontal and not vertical:
                break

            ratios = []

            if horizontal and metrics["bound_width"] > 0:
                ratios.append(
                    metrics["available_width"]
                    / metrics["bound_width"]
                )

            if vertical and metrics["bound_height"] > 0:
                ratios.append(
                    metrics["available_height"]
                    / metrics["bound_height"]
                )

            if not ratios:
                break

            suggested_scale = (
                current_scale
                * min(ratios)
                * 0.97
            )

            new_scale = max(
                min_scale,
                min(current_scale * 0.95, suggested_scale),
            )

            if new_scale >= current_scale - 0.005:
                new_scale = max(
                    min_scale,
                    current_scale * 0.90,
                )

            if new_scale >= current_scale - 0.005:
                break

            _apply_shape_font_scale(
                shape,
                translated_runs,
                new_scale,
            )

            current_scale = new_scale
            shrunk = True

            if current_scale <= min_scale + 0.001:
                break

    horizontal, vertical, metrics = _shape_overflow(shape)

    # Once the font cannot be reduced further, expand the object.
    # Horizontal expansion has priority, then height is recalculated.
    original_lock = _safe_get(
        shape,
        "LockAspectRatio",
        None,
    )

    try:
        if original_lock is not None:
            shape.LockAspectRatio = 0

        if horizontal:
            extra_width = max(
                0.0,
                metrics["bound_width"]
                - metrics["available_width"],
            )

            shape.Width = (
                float(shape.Width)
                + extra_width
                + 2.0
            )

            resized = True
            horizontal, vertical, metrics = _shape_overflow(shape)

        if vertical:
            extra_height = max(
                0.0,
                metrics["bound_height"]
                - metrics["available_height"],
            )

            shape.Height = (
                float(shape.Height)
                + extra_height
                + 2.0
            )

            resized = True

    finally:
        if original_lock is not None:
            try:
                shape.LockAspectRatio = original_lock
            except Exception:
                pass

    return shrunk, resized


def _iter_text_cells(worksheet):
    try:
        cells = worksheet.UsedRange.SpecialCells(
            XL_CELL_TYPE_CONSTANTS,
            XL_TEXT_VALUES,
        )
    except Exception:
        return

    seen = set()

    for cell in cells.Cells:
        try:
            if bool(cell.MergeCells):
                cell = cell.MergeArea.Cells(1, 1)

            address = str(cell.Address)

            if address in seen:
                continue

            seen.add(address)

            value = cell.Value2

            if isinstance(value, str) and value.strip():
                yield cell

        except Exception:
            continue


def _translate_cells(
    worksheet,
    translate_func,
    max_chars,
):
    translated = 0

    for cell in _iter_text_cells(worksheet):
        try:
            text = str(cell.Value2)
            runs = _extract_cell_runs(
                cell,
                text,
            )

            translated_runs = translate_rich_runs(
                runs,
                translate_func=translate_func,
                max_chars=max_chars,
            )

            _write_cell_runs(
                cell,
                translated_runs,
            )

            translated += 1

        except Exception as exc:
            address = _safe_get(
                cell,
                "Address",
                "<unknown cell>",
            )
            print(
                "    WARNING: skipped Excel cell "
                f"{address}: {type(exc).__name__}: {exc}"
            )

    return translated


def _process_shape(
    shape,
    translate_func,
    max_chars,
):
    translated = 0
    shrunk = 0
    resized = 0

    try:
        shape_type = int(shape.Type)
    except Exception:
        shape_type = None

    if shape_type == MSO_GROUP:
        try:
            group_items = shape.GroupItems

            for index in range(
                1,
                int(group_items.Count) + 1,
            ):
                child = group_items.Item(index)

                child_stats = _process_shape(
                    child,
                    translate_func,
                    max_chars,
                )

                translated += child_stats[0]
                shrunk += child_stats[1]
                resized += child_stats[2]

        except Exception:
            pass

        return translated, shrunk, resized

    try:
        # Use the legacy Excel TextFrame interface for text access and rich
        # formatting. Some Excel versions/shapes expose only a subset of the
        # Office TextFrame2/Font2 members.
        text_frame = shape.TextFrame
        text = str(text_frame.Characters().Text or "")

        if not text.strip():
            return translated, shrunk, resized

        runs = _extract_shape_runs(
            shape,
            text,
        )

        translated_runs = translate_rich_runs(
            runs,
            translate_func=translate_func,
            max_chars=max_chars,
        )

        _write_shape_runs(
            shape,
            translated_runs,
        )

        was_shrunk, was_resized = _fit_shape_text(
            shape,
            translated_runs,
        )

        translated += 1
        shrunk += int(was_shrunk)
        resized += int(was_resized)

    except Exception as exc:
        shape_name = _safe_get(
            shape,
            "Name",
            "<unknown shape>",
        )

        print(
            "    WARNING: skipped Excel object "
            f"{shape_name}: {type(exc).__name__}: {exc}"
        )

    return translated, shrunk, resized


def _translate_shapes(
    worksheet,
    translate_func,
    max_chars,
):
    translated = 0
    shrunk = 0
    resized = 0

    shapes = worksheet.Shapes

    for index in range(
        1,
        int(shapes.Count) + 1,
    ):
        stats = _process_shape(
            shapes.Item(index),
            translate_func,
            max_chars,
        )

        translated += stats[0]
        shrunk += stats[1]
        resized += stats[2]

    return translated, shrunk, resized


def translate_excel_file(
    input_path,
    output_path,
    translate_func: Callable[[str], str],
    max_chars=1200,
):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()

    if input_path.suffix.lower() not in EXCEL_EXTENSIONS:
        raise ValueError(
            f"Unsupported Excel file: {input_path.name}"
        )

    if input_path == output_path:
        raise ValueError(
            "Excel input and output paths must be different."
        )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    shutil.copy2(
        input_path,
        output_path,
    )

    try:
        output_path.chmod(
            output_path.stat().st_mode
            | stat.S_IWRITE
        )
    except OSError:
        pass

    pythoncom, win32com_client = _load_excel_com()

    excel = None
    workbook = None

    total_cells = 0
    total_shapes = 0
    total_shrunk = 0
    total_resized = 0

    pythoncom.CoInitialize()

    try:
        excel = win32com_client.DispatchEx(
            "Excel.Application"
        )

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        try:
            # msoAutomationSecurityForceDisable.
            # Prevent Workbook_Open macros from running while preserving them.
            excel.AutomationSecurity = 3
        except Exception:
            pass

        workbook = excel.Workbooks.Open(
            str(output_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        print(f"Excel              : {input_path.name}")

        for index in range(
            1,
            int(workbook.Worksheets.Count) + 1,
        ):
            worksheet = workbook.Worksheets.Item(index)

            print(
                f"  Sheet: {worksheet.Name}"
            )

            cell_count = _translate_cells(
                worksheet,
                translate_func=translate_func,
                max_chars=max_chars,
            )

            (
                shape_count,
                shrunk_count,
                resized_count,
            ) = _translate_shapes(
                worksheet,
                translate_func=translate_func,
                max_chars=max_chars,
            )

            total_cells += cell_count
            total_shapes += shape_count
            total_shrunk += shrunk_count
            total_resized += resized_count

            print(
                f"    cells={cell_count}, "
                f"objects={shape_count}, "
                f"font_shrunk={shrunk_count}, "
                f"objects_resized={resized_count}"
            )

        workbook.Save()

        print(
            "Excel translated    : "
            f"cells={total_cells}, "
            f"objects={total_shapes}, "
            f"font_shrunk={total_shrunk}, "
            f"objects_resized={total_resized}"
        )

    finally:
        if workbook is not None:
            try:
                workbook.Close(
                    SaveChanges=False
                )
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        pythoncom.CoUninitialize()
