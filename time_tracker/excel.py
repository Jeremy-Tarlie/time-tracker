import os
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet

from time_tracker.paths import normalize_excel_path

HEADER_FILL = PatternFill("solid", fgColor="2D6A4F")
SUBHEAD_FILL = PatternFill("solid", fgColor="40916C")
ALT_FILL = PatternFill("solid", fgColor="F0FFF4")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")


def thin_border() -> Border:
    side = Side(style="thin", color="B7E4C7")
    return Border(left=side, right=side, top=side, bottom=side)


def _row_has_data(values: tuple[object, ...]) -> bool:
    return any(value is not None and str(value).strip() != "" for value in values)


def _last_row_with_data(ws: Worksheet, *, min_row: int = 2) -> int:
    """Dernière ligne contenant des données (openpyxl.max_row peut être incorrect après édition manuelle)."""
    last = min_row - 1
    for row_idx, row in enumerate(ws.iter_rows(min_row=min_row, values_only=True), start=min_row):
        if _row_has_data(row):
            last = row_idx
    return last


def _next_session_row(ws: Worksheet) -> int:
    last = _last_row_with_data(ws, min_row=2)
    return max(last + 1, 2)


def _clear_summary_row(ws2: Worksheet, row_idx: int) -> None:
    for col in range(1, 4):
        cell = ws2.cell(row=row_idx, column=col)
        cell.value = None
        cell.fill = WHITE_FILL


def ensure_excel(excel_path: str) -> None:
    if os.path.exists(excel_path):
        return
    parent_dir = os.path.dirname(excel_path)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Sessions")
    else:
        ws.title = "Sessions"
    headers = ["Date", "Projet", "Début", "Fin", "Durée (h)", "Durée (min)", "Notes"]
    col_widths = [14, 22, 10, 10, 12, 14, 30]
    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=i, value=header)
        cell.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border()
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("Résumé par projet")
    ws2["A1"] = "Résumé par Projet"
    ws2["A1"].font = Font(bold=True, color="FFFFFF", name="Calibri", size=14)
    ws2["A1"].fill = HEADER_FILL
    ws2["A1"].alignment = Alignment(horizontal="center")
    ws2.merge_cells("A1:C1")
    ws2.row_dimensions[1].height = 28
    for i, header in enumerate(["Projet", "Total heures", "Nb sessions"], 1):
        cell = ws2.cell(row=2, column=i, value=header)
        cell.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.fill = SUBHEAD_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border()
    ws2.column_dimensions["A"].width = 24
    ws2.column_dimensions["B"].width = 16
    ws2.column_dimensions["C"].width = 14
    ws2.freeze_panes = "A3"
    wb.save(excel_path)


def log_session(
    project: str,
    start_dt: datetime,
    end_dt: datetime,
    notes: str,
    excel_path: str,
) -> None:
    ensure_excel(excel_path)
    wb = load_workbook(excel_path)
    ws = wb["Sessions"]
    duration_h = (end_dt - start_dt).total_seconds() / 3600
    duration_min = (end_dt - start_dt).total_seconds() / 60
    row_idx = _next_session_row(ws)
    fill = ALT_FILL if row_idx % 2 == 0 else WHITE_FILL
    values: list[str | float] = [
        start_dt.strftime("%d/%m/%Y"),
        project,
        start_dt.strftime("%H:%M"),
        end_dt.strftime("%H:%M"),
        round(duration_h, 4),
        round(duration_min, 1),
        notes,
    ]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.font = Font(name="Calibri", size=10)
        cell.fill = fill
        cell.border = thin_border()
        cell.alignment = Alignment(
            horizontal="center" if col != 7 else "left",
            vertical="center",
        )
    ws.cell(row=row_idx, column=5).number_format = "0.00"
    ws.cell(row=row_idx, column=6).number_format = "0.0"
    _refresh_summary(wb)
    wb.save(excel_path)


def rename_project_in_excel(old_name: str, new_name: str, excel_path: str) -> None:
    if not os.path.exists(excel_path):
        return
    wb = load_workbook(excel_path)
    ws = wb["Sessions"]
    for row in ws.iter_rows(min_row=2):
        if row[1].value == old_name:
            row[1].value = new_name
    _refresh_summary(wb)
    wb.save(excel_path)


def _refresh_summary(wb: WorkbookType) -> None:
    ws = wb["Sessions"]
    ws2 = wb["Résumé par projet"]
    totals: dict[str, dict[str, float]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1]:
            project = str(row[1])
            hours = float(row[4] or 0)
            entry = totals.get(project, {"h": 0.0, "n": 0.0})
            entry["h"] += hours
            entry["n"] += 1
            totals[project] = entry
    sorted_projects = sorted(totals.items())
    auto_end = 2 + len(sorted_projects)
    for i in range(3, auto_end + 1):
        _clear_summary_row(ws2, i)
    # Supprimer d'anciennes lignes auto en doublon sans toucher aux lignes manuelles.
    for i in range(auto_end + 1, _last_row_with_data(ws2, min_row=3) + 1):
        name = ws2.cell(row=i, column=1).value
        if name is not None and str(name) in totals:
            _clear_summary_row(ws2, i)
    for i, (project, data) in enumerate(sorted_projects, start=3):
        fill2 = ALT_FILL if i % 2 == 0 else WHITE_FILL
        values: list[str | float] = [project, round(data["h"], 2), int(data["n"])]
        for col, val in enumerate(values, 1):
            cell = ws2.cell(row=i, column=col, value=val)
            cell.font = Font(name="Calibri", size=10)
            cell.fill = fill2
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center" if col != 1 else "left")
        ws2.cell(row=i, column=2).number_format = "0.00"


def open_excel_file(excel_path: str, *, parent: tk.Misc | None = None) -> bool:
    raw = excel_path.strip()
    if not raw:
        messagebox.showwarning("Chemin manquant", "Indiquez un chemin de fichier.", parent=parent)
        return False
    path = normalize_excel_path(raw)
    parent_dir = os.path.dirname(path)
    if parent_dir and not os.path.isdir(parent_dir):
        messagebox.showwarning(
            "Dossier introuvable",
            f"Le dossier suivant n'existe pas :\n{parent_dir}",
            parent=parent,
        )
        return False
    ensure_excel(path)
    try:
        if os.name == "nt":
            os.startfile(path)
        else:
            try:
                sysname = os.uname().sysname
            except AttributeError:
                sysname = ""
            if sysname == "Darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
    except OSError as exc:
        messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier :\n{exc}", parent=parent)
        return False
    return True
