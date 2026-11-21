"""settings.py keeps small preferences in a JSON file."""
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from gradefiller import settings


class SettingsTests(unittest.TestCase):
    def test_remembers_and_survives_a_broken_file(self):
        d = Path(tempfile.mkdtemp())
        with mock.patch("gradefiller.settings.app_data_dir", return_value=d):
            self.assertIsNone(settings.get(settings.NEVER_ASK_TO_SAVE_PASSWORD))
            settings.set(settings.NEVER_ASK_TO_SAVE_PASSWORD, True)
            self.assertTrue(settings.get(settings.NEVER_ASK_TO_SAVE_PASSWORD))
            (d / "settings.json").write_text("{not json")
            self.assertFalse(settings.get(settings.NEVER_ASK_TO_SAVE_PASSWORD, False))
            settings.set("x", 1)  # rewrites the broken file
            self.assertEqual(settings.get("x"), 1)


if __name__ == "__main__":
    unittest.main()
