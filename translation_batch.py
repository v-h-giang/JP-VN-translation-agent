"""Reusable translation deduplication and contextual multi-segment batching.

All helpers are standard-library-only so they can also be imported by
LibreOffice's bundled Python interpreter.
"""
import re
import unicodedata

BATCH_MAX_CHARS = 1100
BATCH_MAX_ITEMS = 16


def translatable(value):
    """Do not send empty/numeric/Japanese-only strings to TranslateGemma."""
    if not value or not value.strip():
        return False
    for char in value:
        if unicodedata.category(char).startswith("L"):
            cp = ord(char)
            if cp <= 0x024F or 0x1E00 <= cp <= 0x1EFF:
                return True
    return False


def _translate_one(value, translate_func):
    """Never exceed the ~1100 source-character budget on one model call."""
    if len(value) <= BATCH_MAX_CHARS:
        return translate_func(value)
    pieces = []
    while value:
        stop = min(len(value), BATCH_MAX_CHARS)
        if stop < len(value):
            cut = max(value.rfind("\n", 0, stop),
                      value.rfind(" ", 0, stop))
            if cut >= 500:
                stop = cut + 1
        pieces.append(translate_func(value[:stop]))
        value = value[stop:]
    return "".join(pieces)


def translate_context_batch(texts, translate_func):
    """Translate unique strings as multi-line context with strict ID mapping.

    If model output loses/reorders IDs or changes the shape of the response,
    fall back to individually translating the affected group.
    """
    if not texts:
        return []

    results = []
    pending = []
    pending_size = 0

    def flush(group):
        if len(group) == 1:
            return [_translate_one(group[0], translate_func)]

        # Model sees multiple relevant segments as context, while IDs retain
        # explicit ownership. Encode internal newlines to keep one segment per
        # physical line; the model must preserve the sequence.
        payload = "\n".join(
            "[[SEG_{:03d}]] {}".format(
                index, text.replace("\\", "\\\\").replace("\n", "\\n")
            )
            for index, text in enumerate(group)
        )
        prompt = (
            "Translate each numbered Vietnamese segment into Japanese. "
            "Keep exactly one line per ID; copy every [[SEG_NNN]] ID unchanged. "
            "Use the other segments for context, but do not merge entries. "
            "Preserve literal \\n sequences for line breaks. "
            "Return translations only, without explanations.\n\n" + payload
        )
        try:
            output = translate_func(prompt)
            matches = re.findall(
                r"(?m)^\s*\[\[SEG_(\d{3})\]\]\s*(.*)$", output
            )
            mapped = {}
            for key, value in matches:
                index = int(key)
                if index >= len(group) or index in mapped:
                    raise ValueError("Unexpected or duplicate segment ID")
                mapped[index] = value
            if len(mapped) != len(group):
                raise ValueError("Missing segment IDs in model response")

            decoded = []
            for index in range(len(group)):
                text = mapped[index].replace("\\n", "\n").replace("\\\\", "\\")
                if not text.strip() and group[index].strip():
                    raise ValueError("Model returned empty translation")
                decoded.append(text)
            return decoded
        except Exception as exc:
            print(
                "WARNING: contextual translation did not preserve IDs "
                f"({type(exc).__name__}: {exc}); translating segments "
                "individually."
            )
            return [_translate_one(text, translate_func) for text in group]

    for value in texts:
        # Oversized source texts get their own bounded requests.
        item_size = len(value) + 24
        if pending and (
            pending_size + item_size > BATCH_MAX_CHARS
            or len(pending) >= BATCH_MAX_ITEMS
        ):
            results.extend(flush(pending))
            pending = []
            pending_size = 0
        pending.append(value)
        pending_size += item_size

    if pending:
        results.extend(flush(pending))
    return results
