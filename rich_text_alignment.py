"""Conservative mapping from UNO text portions to the exact cell string.

LibreOffice paragraph enumerations may exclude separators. We may add/remove
whitespace, but never invent/erase other characters or silently flatten styles.
"""
from difflib import SequenceMatcher


class RichTextAlignmentError(ValueError):
    """Source text cannot be safely reconstructed with existing rich styles."""


def _separator(text):
    return all(ch.isspace() or ch in ("\x00", "\x0b", "\x0c") for ch in text)


def _merge(parts):
    result = []
    for item in parts:
        value = item["text"]
        if not value:
            continue
        style = item["style"]
        if result and result[-1]["style"] == style:
            result[-1]["text"] += value
        else:
            result.append({"text": value, "style": dict(style)})
    return result


def align_rich_text_runs(full, runs, max_diff_chars=12000):
    """Return styled runs that reproduce the entire original text exactly.

    Text absent from UNO portions must be only whitespace/newline separators;
    gaps inherit neighboring styles. Non-whitespace mismatches are unsafe.
    """
    if not isinstance(full, str):
        raise TypeError("full must be str")
    if not full:
        return []
    if not runs:
        raise RichTextAlignmentError("No text portions were exposed")
    if any(not isinstance(run.get("text"), str)
           or not isinstance(run.get("style"), dict) for run in runs):
        raise RichTextAlignmentError("Malformed UNO rich text portions")

    merged = _merge(runs)
    raw = "".join(x["text"] for x in merged)
    if raw == full:
        return merged

    # Fast path: UNO simply left newline separators out of portions.
    recovered = []
    offset = 0
    valid_fast_path = True
    for run in merged:
        fragment = run["text"]
        pos = full.find(fragment, offset)
        if pos < offset or not _separator(full[offset:pos]):
            valid_fast_path = False
            break
        if pos > offset:
            inherited = recovered[-1]["style"] if recovered else run["style"]
            recovered.append({"text": full[offset:pos], "style": inherited})
        recovered.append(run)
        offset = pos + len(fragment)
    if valid_fast_path and _separator(full[offset:]):
        if offset < len(full):
            recovered.append({"text": full[offset:], "style": recovered[-1]["style"]})
        return _merge(recovered)

    if max(len(raw), len(full)) > max_diff_chars:
        raise RichTextAlignmentError(
            "Ambiguous large rich-text cell; refused unsafe alignment"
        )

    # Handle different newline representations or repeated substrings.
    styles = []
    for run in merged:
        styles.extend([run["style"]] * len(run["text"]))
    assigned = [None] * len(full)
    matcher = SequenceMatcher(a=raw, b=full, autojunk=False)
    for op, a0, a1, b0, b1 in matcher.get_opcodes():
        if op == "equal":
            assigned[b0:b1] = styles[a0:a1]
        elif op == "insert":
            if not _separator(full[b0:b1]):
                raise RichTextAlignmentError("Non-whitespace text missing from UNO portions")
        elif op == "delete":
            if not _separator(raw[a0:a1]):
                raise RichTextAlignmentError("UNO portions contain non-source text")
        elif op == "replace":
            if not (_separator(raw[a0:a1]) and _separator(full[b0:b1])):
                raise RichTextAlignmentError("UNO portion text conflicts with cell text")

    # Inherit the previous run's style for newly found whitespace.
    last_style = None
    for index, style in enumerate(assigned):
        if style is not None:
            last_style = style
        elif last_style is not None:
            assigned[index] = last_style
    next_style = next((x for x in assigned if x is not None), merged[0]["style"])
    for index, style in enumerate(assigned):
        if style is None:
            assigned[index] = next_style
        else:
            next_style = style

    aligned = _merge([
        {"text": char, "style": assigned[index]}
        for index, char in enumerate(full)
    ])
    if "".join(x["text"] for x in aligned) != full:
        raise RichTextAlignmentError("Rich-text reconstruction did not match original")
    return aligned
