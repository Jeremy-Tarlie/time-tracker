"""Gardes d'écriture Excel : format OOXML et verrous."""

from __future__ import annotations

import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from main import (
    ExcelWriteError,
    _assert_excel_ooxml,
    _excel_lock_paths,
    ensure_excel,
    format_session_duration,
    log_session,
)


class TestFormatSessionDuration(unittest.TestCase):
    def test_seconds_only(self) -> None:
        self.assertEqual(format_session_duration(5), "5s")

    def test_minutes_and_seconds(self) -> None:
        self.assertEqual(format_session_duration(65), "1min 05s")

    def test_hours_minutes_seconds(self) -> None:
        self.assertEqual(format_session_duration(3661), "1h 01min 01s")



class TestExcelWriteGuard(unittest.TestCase):
    def test_lock_paths_include_excel_and_libreoffice(self) -> None:
        path = r"C:\data\heures_projet.xlsx"
        locks = _excel_lock_paths(path)
        self.assertIn(r"C:\data\~$heures_projet.xlsx", locks)
        self.assertIn(r"C:\data\.~lock.heures_projet.xlsx#", locks)

    def test_assert_ooxml_rejects_ole_compound_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            fake = Path(tmp) / "fake.xlsx"
            fake.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 32)
            with self.assertRaises(ExcelWriteError) as ctx:
                _assert_excel_ooxml(str(fake))
            self.assertIn("vrai .xlsx", str(ctx.exception))

    def test_assert_ooxml_accepts_zip_magic(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "ok.xlsx"
            ensure_excel(str(path))
            _assert_excel_ooxml(str(path))  # ne lève pas

    def test_log_session_rejects_ole_file_without_writing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            fake = Path(tmp) / "heures_projet.xlsx"
            fake.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64)
            before = fake.read_bytes()
            start = datetime(2026, 10, 21, 15, 0, 0)
            end = datetime(2026, 10, 21, 15, 0, 3)
            with self.assertRaises(ExcelWriteError):
                log_session("Projet 1", start, end, "note", str(fake))
            self.assertEqual(fake.read_bytes(), before)

    def test_log_session_appends_on_valid_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "heures_projet.xlsx")
            ensure_excel(path)
            start = datetime(2026, 10, 21, 15, 0, 0)
            end = datetime(2026, 10, 21, 15, 5, 0)
            row = log_session("Projet 1", start, end, "ok", path)
            self.assertEqual(row, 2)
            from openpyxl import load_workbook
            wb = load_workbook(path)
            ws = wb["Sessions"]
            self.assertEqual(ws.cell(2, 2).value, "Projet 1")
            self.assertEqual(ws.cell(2, 7).value, "ok")
            wb.close()


if __name__ == "__main__":
    unittest.main()
