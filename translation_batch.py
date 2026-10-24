"""Pure-Python contextual spreadsheet translation utilities.

No UNO/Excel dependencies: imported by both backends and unit-testable.
"""
import json
import re
import unicodedata

BATCH_MAX_CHARS = 1100
BATCH_MAX_ITEMS = 16
# Bump when translation/cache semantics change.
BATCH_ALGORITHM_VERSION = 3

VIETNAMESE_ACCENTS = re.compile(
    r"[À-ỹĐđĂăÂâÊêÔôƠơƯư]", re.UNICODE
)
JAPANESE = re.compile(r"[\u3040-\u30ff\u3400-\u9fff]")
LATIN = re.compile(r"[A-Za-z\u00c0-\u024f\u1e00-\u1eff]")


def _literal_pattern(term):
    # Respect token boundaries: ID shouldn't match INSIDE or GRID.
    boundary_start = r"(?<![A-Za-z0-9_])" if term[0].isalnum() else ""
    boundary_end = r"(?![A-Za-z0-9_])" if term[-1].isalnum() else ""
    return re.compile(boundary_start + re.escape(term) + boundary_end)


def literal_terms_from_glossary(glossary):
    """Identity entries in glossary.json mean 'never translate this token'."""
    return tuple(
        sorted(
            (key for key, value in glossary.items()
             if key and key == value and re.search("[A-Za-z]", key)),
            key=len,
            reverse=True,
        )
    )


def translatable(value, literal_terms=()):
    """Skip Japanese + Latin technical identifiers unless Vietnamese is present."""
    if not value or not value.strip():
        return False
    text = unicodedata.normalize("NFC", value)
    if not LATIN.search(text):
        return False
    if JAPANESE.search(text) and not VIETNAMESE_ACCENTS.search(text):
        return False
    if any(_literal_pattern(term).fullmatch(text.strip()) for term in literal_terms):
        return False
    return True


def _literal_counts(text, terms):
    return {
        term: len(_literal_pattern(term).findall(text))
        for term in terms if _literal_pattern(term).search(text)
    }


def _literals_ok(source, target, terms):
    return all(
        len(_literal_pattern(term).findall(target)) == count
        for term, count in _literal_counts(source, terms).items()
    )


def _split_preserving_literals(source, terms):
    matches = []
    for term in terms:
        for match in _literal_pattern(term).finditer(source):
            matches.append((match.start(), match.end(), term))
    matches.sort(key=lambda item: (item[0], -(item[1] - item[0])))
    selected = []
    end = 0
    for start, finish, term in matches:
        if start < end:
            continue
        selected.append((start, finish, term))
        end = finish
    return selected


def _translate_one(value, translate_func, literal_terms=()):
    """Bounded requests; if a glossary identity token is lost, translate around it."""
    def plain(text):
        if not text or not text.strip() or not translatable(text, literal_terms):
            return text
        if len(text) <= BATCH_MAX_CHARS:
            translated = translate_func(text)
            if not isinstance(translated, str) or not translated.strip():
                raise RuntimeError("Empty translation for nonempty segment")
            return translated

        chunks = []
        remainder = text
        while remainder:
            limit = min(len(remainder), BATCH_MAX_CHARS)
            if limit < len(remainder):
                cut = max(
                    remainder.rfind("\n", 0, limit),
                    remainder.rfind(" ", 0, limit),
                )
                if cut >= 500:
                    limit = cut + 1
            chunks.append(plain(remainder[:limit]))
            remainder = remainder[limit:]
        return "".join(chunks)

    candidate = plain(value)
    if _literals_ok(value, candidate, literal_terms):
        return candidate

    # Deterministic repair: never trust the model to restore acronym spelling.
    # Translate spans between the original literal terms separately.
    preserved = _split_preserving_literals(value, literal_terms)
    if not preserved:
        raise RuntimeError("Translation unexpectedly altered a protected glossary term")
    pieces = []
    offset = 0
    for begin, end, term in preserved:
        pieces.append(plain(value[offset:begin]))
        pieces.append(value[begin:end])
        offset = end
    pieces.append(plain(value[offset:]))
    result = "".join(pieces)
    if not _literals_ok(value, result, literal_terms):
        raise RuntimeError("Protected glossary tokens could not be preserved")
    return result


def translate_context_batch(texts, translate_func, literal_terms=()):
    """Translate text batches with strict 1:1 mapping and lossless escaping.

    JSON-string payloads handle backslashes, embedded newlines and quotes.
    When the model fails the protocol, translate each original string.
    """
    if not texts:
        return []
    results = []
    pending = []
    size = 0

    def flush(group):
        if len(group) == 1:
            return [_translate_one(group[0], translate_func, literal_terms)]

        lines = [
            "[[SEG_{:03d}]] {}".format(
                index, json.dumps(text, ensure_ascii=False),
            )
            for index, text in enumerate(group)
        ]
        prompt = (
            "Translate the Vietnamese text inside each JSON string into "
            "Japanese. Return exactly one line for every ID in order: "
            "[[SEG_NNN]] followed by a valid JSON string. Preserve ID values, "
            "technical names, backslashes, escapes and JSON syntax. "
            "No explanations.\n\n" + "\n".join(lines)
        )
        try:
            output = translate_func(prompt)
            matches = re.findall(
                r"(?m)^\s*\[\[SEG_(\d{3})\]\]\s*(.+?)\s*$", output
            )
            converted = {}
            for index_str, payload in matches:
                index = int(index_str)
                if index in converted or index >= len(group):
                    raise ValueError("Duplicate/invalid segment ID")
                translated = json.loads(payload)
                if not isinstance(translated, str) or not translated.strip():
                    raise ValueError("Segment returned empty/non-string translation")
                if not _literals_ok(group[index], translated, literal_terms):
                    raise ValueError("Glossary literal was changed")
                converted[index] = translated
            if len(converted) != len(group):
                raise ValueError("Missing segment IDs")
            return [converted[i] for i in range(len(group))]
        except (ValueError, TypeError, RuntimeError, json.JSONDecodeError) as exc:
            print(
                "WARNING: contextual translation ID/JSON mismatch "
                f"({type(exc).__name__}: {exc}); retrying individually."
            )
            return [
                _translate_one(value, translate_func, literal_terms)
                for value in group
            ]

    for value in texts:
        estimate = len(json.dumps(value, ensure_ascii=False)) + 24
        if pending and (size + estimate > BATCH_MAX_CHARS
                        or len(pending) >= BATCH_MAX_ITEMS):
            results.extend(flush(pending))
            pending = []
            size = 0
        pending.append(value)
        size += estimate
    if pending:
        results.extend(flush(pending))
    return results
