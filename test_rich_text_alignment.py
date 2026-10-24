import unittest
from rich_text_alignment import align_rich_text_runs, RichTextAlignmentError


class RichTextAlignmentTests(unittest.TestCase):
    def test_mixed_style_multiline_strikeout(self):
        runs = [
            {"text": "旧仕様", "style": {"CharStrikeout": 1, "CharColor": 255}},
            {"text": "新仕様", "style": {"CharWeight": 150, "CharHeight": 8}},
        ]
        restored = align_rich_text_runs("旧仕様\n新仕様", runs)
        self.assertEqual("".join(x["text"] for x in restored), "旧仕様\n新仕様")
        self.assertEqual(restored[0]["style"]["CharStrikeout"], 1)
        self.assertEqual(restored[-1]["style"]["CharHeight"], 8)

    def test_repeated_substrings(self):
        runs = [{"text": "abc", "style": {"i": 1}},
                {"text": "abc", "style": {"i": 2}}]
        restored = align_rich_text_runs("abc\nabc", runs)
        self.assertEqual(restored[-1]["style"]["i"], 2)

    def test_trailing_newline(self):
        restored = align_rich_text_runs("abc\n", [
            {"text": "abc", "style": {"strike": True}},
        ])
        self.assertEqual(restored[0]["text"], "abc\n")

    def test_newline_normalization(self):
        restored = align_rich_text_runs("a\nb", [
            {"text": "a\rb", "style": {"size": 12}},
        ])
        self.assertEqual(restored[0]["text"], "a\nb")

    def test_missing_visible_text_must_not_be_ignored(self):
        with self.assertRaises(RichTextAlignmentError):
            align_rich_text_runs("旧仕様\n重要\n新仕様", [
                {"text": "旧仕様", "style": {}},
                {"text": "新仕様", "style": {}},
            ])

    def test_no_runs_fails_safe(self):
        with self.assertRaises(RichTextAlignmentError):
            align_rich_text_runs("abc", [])


if __name__ == "__main__":
    unittest.main()
