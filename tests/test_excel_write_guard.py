"""Gardes d'écriture Excel : format OOXML et verrous."""

from __future__ import annotations

import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from time_tracker.excel import (
    ExcelWriteError,
    _assert_excel_ooxml,
    _excel_lock_paths,
    ensure_excel,
    format_session_duration,
    log_session,
    rename_project_in_excel,
    sanitize_excel_text,
)


class TestFormatSessionDuration(unittest.TestCase):
    def test_seconds_only(self) -> None:
        self.assertEqual(format_session_duration(5), "5s")

    def test_minutes_and_seconds(self) -> None:
        self.assertEqual(format_session_duration(65), "1min 05s")

    def test_hours_minutes_seconds(self) -> None:
        self.assertEqual(format_session_duration(3661), "1h 01min 01s")


class TestSanitizeExcelText(unittest.TestCase):
    def test_prefixes_formula_injection(self) -> None:
        self.assertEqual(sanitize_excel_text("=cmd|'/C calc'!A0"), "'=cmd|'/C calc'!A0")
        self.assertEqual(sanitize_excel_text("+1+1"), "'+1+1")
        self.assertEqual(sanitize_excel_text("-1"), "'-1")
        self.assertEqual(sanitize_excel_text("@SUM(A1)"), "'@SUM(A1)")

    def test_strips_leading_controls_before_formula_check(self) -> None:
        self.assertEqual(sanitize_excel_text("\t=1+1"), "'=1+1")
        self.assertEqual(sanitize_excel_text(" =1+1"), "'=1+1")
        self.assertEqual(sanitize_excel_text("\ufeff=1+1"), "'=1+1")
        self.assertEqual(sanitize_excel_text("\u200b=1+1"), "'=1+1")
        self.assertEqual(sanitize_excel_text("\r\n@SUM(A1)"), "'@SUM(A1)")

    def test_leaves_normal_text(self) -> None:
        self.assertEqual(sanitize_excel_text("note normale"), "note normale")
        self.assertEqual(sanitize_excel_text(""), "")


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
            _assert_excel_ooxml(str(path))

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

    def test_log_session_appends_after_manual_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "heures_projet.xlsx")
            ensure_excel(path)
            start = datetime(2026, 10, 21, 15, 0, 0)
            end = datetime(2026, 10, 21, 15, 5, 0)
            self.assertEqual(log_session("A", start, end, "app1", path), 2)

            from openpyxl import load_workbook

            wb = load_workbook(path)
            ws = wb["Sessions"]
            ws.cell(3, 1).value = "01/01/2026"
            ws.cell(3, 2).value = "Manuel"
            ws.cell(5, 1).value = "02/01/2026"
            ws.cell(5, 2).value = "Manuel2"
            wb.save(path)
            wb.close()

            row = log_session(
                "B",
                datetime(2026, 1, 2, 10, 0, 0),
                datetime(2026, 1, 2, 10, 0, 5),
                "app2",
                path,
            )
            self.assertEqual(row, 6)
            wb = load_workbook(path)
            ws = wb["Sessions"]
            self.assertEqual(ws.cell(3, 2).value, "Manuel")
            self.assertEqual(ws.cell(5, 2).value, "Manuel2")
            self.assertEqual(ws.cell(6, 2).value, "B")
            wb.close()

    def test_log_session_sanitizes_formula_notes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "heures_projet.xlsx")
            ensure_excel(path)
            start = datetime(2026, 10, 21, 15, 0, 0)
            end = datetime(2026, 10, 21, 15, 0, 5)
            log_session("Projet 1", start, end, "=HYPERLINK(\"http://x\")", path)
            from openpyxl import load_workbook

            wb = load_workbook(path)
            self.assertEqual(wb["Sessions"].cell(2, 7).value, "'=HYPERLINK(\"http://x\")")
            wb.close()

    def test_rename_raises_on_ole_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            fake = Path(tmp) / "heures_projet.xlsx"
            fake.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64)
            with self.assertRaises(ExcelWriteError):
                rename_project_in_excel("A", "B", str(fake))

    def test_refresh_summary_sanitizes_legacy_formula_project(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "heures_projet.xlsx")
            ensure_excel(path)
            from openpyxl import load_workbook
            from time_tracker.excel import _refresh_summary, _save_workbook_atomic

            wb = load_workbook(path)
            ws = wb["Sessions"]
            ws.cell(2, 1).value = "21/10/2026"
            ws.cell(2, 2).value = "=1+1"
            ws.cell(2, 5).value = 0.5
            _refresh_summary(wb)
            _save_workbook_atomic(wb, path)
            wb.close()

            wb = load_workbook(path)
            summary_name = wb["Résumé par projet"].cell(3, 1).value
            self.assertEqual(summary_name, "'=1+1")
            self.assertEqual(wb["Résumé par projet"].cell(3, 1).data_type, "s")
            wb.close()


if __name__ == "__main__":
    unittest.main()
