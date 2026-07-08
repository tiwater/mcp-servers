import unittest

from tiwater_pdf.cli import _extract_json_object


class ExtractJsonObjectTest(unittest.TestCase):
    def test_accepts_plain_object(self):
        self.assertEqual(_extract_json_object('{"text":"a","tables":[],"warnings":[]}')["text"], "a")

    def test_accepts_json_fence(self):
        data = _extract_json_object('```json\n{"text":"a","tables":[],"warnings":[]}\n```')
        self.assertEqual(data["text"], "a")

    def test_uses_first_complete_object_when_model_adds_extra_json(self):
        data = _extract_json_object('{"text":"a","tables":[],"warnings":[]}\n{"debug":"ignored"}')
        self.assertEqual(data["text"], "a")

    def test_uses_first_complete_object_when_model_adds_trailing_text(self):
        data = _extract_json_object('{"text":"a","tables":[],"warnings":[]}\nDone.')
        self.assertEqual(data["text"], "a")

    def test_rejects_non_object_json(self):
        with self.assertRaises(ValueError):
            _extract_json_object('[{"text":"a"}]')


if __name__ == "__main__":
    unittest.main()
