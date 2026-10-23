"""LibreOffice UNO worker. Run with Python that can import LibreOffice's 'uno'.

Communication with the parent translator is newline-delimited JSON via stdio.
The parent owns Ollama, so the UNO worker does not import third-party modules.
"""
import json
import sys
import time
import traceback
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue
# CellContentType is defined in com.sun.star.table (not .sheet).
# Avoid importing UNO enum members as Python modules: pyuno dynamically
# resolves them, and a wrong namespace prevents the worker from starting.
CELL_TEXT = uno.Enum("com.sun.star.table.CellContentType", "TEXT")


MIN_FONT_PT = 8.0
MAX_RUNS = 300
FORMAT_PROPS = (
    "CharFontName", "CharHeight", "CharWeight", "CharPosture",
    "CharColor", "CharUnderline", "CharStrikeout", "CharEscapement",
)


def send(kind, **values):
    sys.stdout.write(json.dumps({"type": kind, **values}, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def receive():
    line = sys.stdin.readline()
    if not line:
        raise RuntimeError("Parent translator closed the pipe")
    return json.loads(line)


def request_translation(runs):
    send("translate", runs=runs)
    reply = receive()
    if not reply.get("ok"):
        raise RuntimeError(str(reply.get("error", "translation failed")))
    return reply["runs"]


def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def get_prop(obj, name):
    try:
        return getattr(obj, name)
    except Exception:
        return None


def get_runs(text_obj):
    # LibreOffice text cursor uses Unicode codepoint offsets as reported by
    # the UNO text range. Group contiguous text portions by font properties.
    runs = []

    def collect(container):
        paragraphs = container.createEnumeration()
        while paragraphs.hasMoreElements():
            paragraph = paragraphs.nextElement()
            try:
                parts = paragraph.createEnumeration()
            except Exception:
                continue
            while parts.hasMoreElements():
                part = parts.nextElement()
                value = part.getString()
                if not value:
                    continue
                style = {}
                for key in FORMAT_PROPS:
                    value_prop = get_prop(part, key)
                    if value_prop is not None:
                        try:
                            json.dumps(value_prop)
                            style[key] = value_prop
                        except TypeError:
                            style[key] = str(value_prop)
                if runs and runs[-1]["style"] == style:
                    runs[-1]["text"] += value
                else:
                    runs.append({"text": value, "style": style})

    collect(text_obj)
    full = text_obj.getString()

    # If portions did not include paragraph separators, prefer a single run:
    # preserving full text is more important than guessing style positions.
    if "".join(r["text"] for r in runs) != full or len(runs) > MAX_RUNS:
        cursor = text_obj.createTextCursor()
        style = {p: get_prop(cursor, p) for p in FORMAT_PROPS}
        style = {k: v for k, v in style.items() if isinstance(v, (int, float, str, bool))}
        return [{"text": full, "style": style}]
    return runs or [{"text": full, "style": {}}]


def set_style(cursor, style):
    for key, val in style.items():
        if key in FORMAT_PROPS:
            try:
                setattr(cursor, key, val)
            except Exception:
                pass


def replace_text(text_obj, translated_runs):
    translated = "".join(run["text"] for run in translated_runs)
    text_obj.setString(translated)
    if len(translated_runs) == 1:
        # Cell-level formatting remains unless explicit new content inherits
        # differently. Force the original font style onto the new text.
        cursor = text_obj.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        set_style(cursor, translated_runs[0]["style"])
        return

    cursor = text_obj.createTextCursor()
    cursor.gotoStart(False)
    for run in translated_runs:
        size = len(run["text"])
        if not size:
            continue
        cursor.goRight(size, True)
        set_style(cursor, run["style"])
        cursor.collapseToEnd()


def cell_text(cell):
    try:
        cell_type = cell.getType()
        # Compare the enum's .value as well, for pyuno versions where
        # different enum proxy instances do not compare identically.
        if cell_type != CELL_TEXT and getattr(cell_type, "value", None) != "TEXT":
            return ""
        return cell.getString()
    except Exception:
        return ""


def fit_shape(shape):
    try:
        if not hasattr(shape, "getSize"):
            return "none"
        text = shape.getString()
        if not text.strip():
            return "none"
        size = shape.getSize()
        original_width = size.Width
        original_height = size.Height
        cursor = shape.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        base_font = get_prop(cursor, "CharHeight") or 10
        base_font = float(base_font)
        # Ask LibreOffice to keep text within the shape when possible.
        for option, value in (("TextAutoGrowWidth", False),
                              ("TextAutoGrowHeight", False),
                              ("TextFitToSize", 0)):
            try:
                setattr(shape, option, value)
            except Exception:
                pass
        # Real-time text metrics for arbitrary UNO drawing shapes are not
        # universally exposed. Use approximate Japanese glyph widths and
        # newline-sensitive height, adding a safety margin.
        vertical_lines = str(text).split("\n")
        font = max(MIN_FONT_PT, base_font)
        line_width = max((sum(1.0 if ord(ch) > 255 else 0.6 for ch in line)
                          for line in vertical_lines), default=0)
        margin = 1.10
        width_at_font = line_width * font * 35.28 * margin
        height_at_font = len(vertical_lines) * font * 35.28 * 1.5
        scale = min(1.0,
                    original_width / max(1, width_at_font),
                    original_height / max(1, height_at_font))
        new_font = max(MIN_FONT_PT, font * scale)
        if new_font < font:
            try:
                cursor.CharHeight = float(new_font)
                font = new_font
            except Exception:
                pass
        required_width = int(line_width * font * 35.28 * margin)
        required_height = int(len(vertical_lines) * font * 35.28 * 1.5)
        resized = False
        if required_width > original_width:
            size.Width = required_width
            resized = True
        if required_height > original_height:
            size.Height = required_height
            resized = True
        if resized:
            shape.setSize(size)
            return "resized"
        if new_font < base_font:
            return "shrunk"
        return "unchanged"
    except Exception as exc:
        send("log", text=f"Shape sizing unavailable: {exc}")
        return "unavailable"


def translate_shape(shape, stats):
    try:
        if shape.supportsService("com.sun.star.drawing.GroupShape"):
            count = shape.getCount()
            for i in range(count):
                translate_shape(shape.getByIndex(i), stats)
            return
        if not hasattr(shape, "getString") or not hasattr(shape, "createTextCursor"):
            return
        original = shape.getString()
        if not original.strip():
            return
        runs = get_runs(shape)
        result = request_translation(runs)
        replace_text(shape, result)
        outcome = fit_shape(shape)
        stats["objects"] += 1
        if outcome == "resized":
            stats["resized"] += 1
        elif outcome == "shrunk":
            stats["shrunk"] += 1
    except Exception as exc:
        stats["errors"] += 1
        send("log", text=f"Skipped shape: {type(exc).__name__}: {exc}")


def translate_sheet(sheet, stats):
    cursor = sheet.createCursor()
    cursor.gotoStartOfUsedArea(False)
    cursor.gotoEndOfUsedArea(True)
    address = cursor.getRangeAddress()
    row_span = address.EndRow - address.StartRow + 1
    col_span = address.EndColumn - address.StartColumn + 1

    if row_span * col_span > 150000:
        raise RuntimeError(
            f"Sheet {sheet.Name} used range too large for safe cell iteration "
            f"({row_span} x {col_span})"
        )
    for row in range(address.StartRow, address.EndRow + 1):
        for col in range(address.StartColumn, address.EndColumn + 1):
            cell = sheet.getCellByPosition(col, row)
            original = cell_text(cell)
            if not original.strip():
                continue
            try:
                runs = get_runs(cell)
                translated_runs = request_translation(runs)
                replace_text(cell, translated_runs)
                stats["cells"] += 1
            except Exception as exc:
                stats["errors"] += 1
                send("log", text=f"Skipped cell {sheet.Name} ({col+1},{row+1}): "
                     f"{type(exc).__name__}: {exc}")

    draw = sheet.getDrawPage()
    for index in range(draw.getCount()):
        translate_shape(draw.getByIndex(index), stats)


def connect(port):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    url = (
        f"uno:socket,host=127.0.0.1,port={port};urp;"
        "StarOffice.ComponentContext"
    )
    last_error = None
    for _ in range(80):
        try:
            context = resolver.resolve(url)
            manager = context.ServiceManager
            desktop = manager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", context
            )
            return desktop
        except Exception as exc:
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"UNO connection failed: {last_error}")


def main():
    port = int(sys.argv[1])
    source = Path(sys.argv[2]).resolve()
    output = Path(sys.argv[3]).resolve()
    desktop = connect(port)
    document = None
    try:
        source_url = uno.systemPathToFileUrl(str(source))
        document = desktop.loadComponentFromURL(
            source_url, "_blank", 0,
            (prop("Hidden", True), prop("ReadOnly", True),
             prop("UpdateDocMode", 0), prop("MacroExecutionMode", 4)),
        )
        if document is None:
            raise RuntimeError(f"LibreOffice cannot open: {source}")
        stats = {"cells": 0, "objects": 0, "shrunk": 0,
                 "resized": 0, "errors": 0}
        for index in range(document.Sheets.getCount()):
            sheet = document.Sheets.getByIndex(index)
            send("log", text=f"Sheet: {sheet.Name}")
            translate_sheet(sheet, stats)

        # Only write output when translation is finished. The source never
        # changes; a failed run does not overwrite an existing output.
        temporary_output = output.parent / (
            output.stem + ".partial" + output.suffix
        )
        filters = {".xlsx": "Calc MS Excel 2007 XML",
                   ".xls": "MS Excel 97"}
        filter_name = filters[source.suffix.lower()]
        document.storeAsURL(
            uno.systemPathToFileUrl(str(temporary_output)),
            (prop("FilterName", filter_name), prop("Overwrite", True)),
        )
        document.close(True)
        document = None
        temporary_output.replace(output)
        send("done", stats=stats)
    finally:
        if document is not None:
            try:
                document.close(True)
            except Exception:
                document.dispose()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        send("error", text=traceback.format_exc())
        sys.exit(1)
