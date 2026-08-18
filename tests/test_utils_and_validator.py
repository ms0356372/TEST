import tempfile
import unittest
from pathlib import Path

from core.exceptions import AnalysisError
from core.utils import calculate_age, safe_number, unique_output_path
from core.validator import validate_excel_input, validate_output_directory


class UtilityTests(unittest.TestCase):
    def test_safe_number_accepts_supported_values(self):
        for value, expected in [(140, 140.0), (140.0, 140.0), ("140", 140.0), ("140 ", 140.0), ("1,400", 1400.0)]:
            with self.subTest(value=value):
                self.assertEqual(safe_number(value), expected)

    def test_safe_number_rejects_invalid_values(self):
        for value in (None, "", "-", "N/A", "未檢", "abc", True, float("nan")):
            with self.subTest(value=value):
                self.assertIsNone(safe_number(value))

    def test_calculate_age_uses_first_four_characters(self):
        for value in ("1987", "19870520", "1987/05/20", "1987-05-20"):
            with self.subTest(value=value):
                self.assertEqual(calculate_age(value, current_year=2026), 39)
        for value in (None, "", "987", "abcd-01-01", "1800", "2099"):
            with self.subTest(value=value):
                self.assertIsNone(calculate_age(value, current_year=2026))

    def test_path_validation_and_unique_output(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "總表.xlsx"
            source.write_bytes(b"test")
            self.assertEqual(validate_excel_input(source, "總表"), source.resolve())
            self.assertEqual(validate_output_directory(root), root.resolve())
            first = unique_output_path(root, "分析")
            first.touch()
            self.assertNotEqual(unique_output_path(root, "分析"), first)

    def test_invalid_paths_raise_user_facing_error(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            with self.assertRaises(AnalysisError):
                validate_excel_input(root / "missing.xlsx", "總表")
            wrong = root / "source.xls"
            wrong.touch()
            with self.assertRaises(AnalysisError):
                validate_excel_input(wrong, "總表")
            with self.assertRaises(AnalysisError):
                validate_output_directory(root / "missing")


if __name__ == "__main__":
    unittest.main()
