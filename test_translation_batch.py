"""Regression tests for lossless spreadsheet batching (stdlib unittest)."""
import unittest

from translation_batch import (
    _translate_one,
    literal_terms_from_glossary,
    translatable,
    translate_context_batch,
)


class SpreadsheetTranslationTests(unittest.TestCase):
    def test_skip_japanese_with_latin_identifiers(self):
        self.assertFalse(translatable("APIの仕様を確認する"))
        self.assertFalse(translatable("WBS更新 / ID管理"))
        self.assertTrue(translatable("cập nhật WBS"))
        self.assertTrue(translatable("Phương châm"))

    def test_skip_standalone_glossary_term(self):
        calls = []
        def translate(text):
            calls.append(text)
            return "訳:" + text
        terms = literal_terms_from_glossary({"API": "API", "WBS": "WBS"})
        self.assertEqual(_translate_one("API", translate, terms), "API")
        self.assertFalse(calls)

    def test_preserve_glossary_term_when_model_changes_it(self):
        calls = []
        def translate(text):
            calls.append(text)
            if text == "cập nhật API":
                return "APIを更新"  # token still present: valid
            return "更新"
        terms = literal_terms_from_glossary({"API": "API"})
        self.assertEqual(_translate_one("cập nhật API", translate, terms), "APIを更新")

    def test_repair_lost_literal(self):
        def translate(text):
            if text == "cập nhật API":
                return "接続を更新"
            return "更新"
        terms = literal_terms_from_glossary({"API": "API"})
        self.assertEqual(_translate_one("cập nhật API", translate, terms), "更新API")

    def test_context_batch_backslashes_and_newlines(self):
        source = [r"C:\new\test", "dòng 1\ndòng 2"]
        def translator(prompt):
            # Mock a compliant model response, including JSON escape syntax.
            self.assertIn('[[SEG_000]] "C:\\\\new\\\\test"', prompt)
            return '[[SEG_000]] "C:\\\\new\\\\test"\n[[SEG_001]] "行1\\n行2"'
        # The second mock is a literal newline inside JSON: protocol must
        # reject it, then retry individual segments; cover with separate test.
        def compliant(_):
            return '[[SEG_000]] "C:\\\\new\\\\test"\n[[SEG_001]] "行1\\u000a行2"'
        self.assertEqual(
            translate_context_batch(source, compliant),
            [source[0], "行1\n行2"],
        )

    def test_malformed_batch_falls_back_safely(self):
        original = ["xác nhận", "cập nhật"]
        seen = []
        def translator(text):
            seen.append(text)
            if "[[SEG_000]]" in text:
                return "モデルがIDを削除した"
            return {"xác nhận": "確認", "cập nhật": "更新"}[text]
        self.assertEqual(
            translate_context_batch(original, translator),
            ["確認", "更新"],
        )
        self.assertEqual(len(seen), 3)

    def test_acronym_boundary(self):
        calls = []
        def translator(text):
            calls.append(text)
            return "更新"
        terms = literal_terms_from_glossary({"ID": "ID"})
        self.assertFalse(translatable("ID", terms))
        # The literal ID must not match inside GRID.
        self.assertEqual(_translate_one("GRID", translator, terms), "更新")


if __name__ == "__main__":
    unittest.main()
