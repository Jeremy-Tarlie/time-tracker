import os
import subprocess
import sys
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import PurePath, PurePosixPath, PureWindowsPath
from tkinter import messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet

from time_tracker.constants import SessionRow
from time_tracker.paths import normalize_excel_path

_FORMULA_PREFIXES = ("=", "+", "-", "@")
# Espaces / contrôles souvent utilisés pour contourner un check sur le 1er caractère.
_FORMULA_LEADING_STRIP = "\t\r\n \ufeff\u200b\u200c\u200d\u2060\u00a0"

HEADER_FILL = PatternFill("solid", fgColor="2D6A4F")
SUBHEAD_FILL = PatternFill("solid", fgColor="40916C")
ALT_FILL = PatternFill("solid", fgColor="F0FFF4")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")

OOXML_ZIP_MAGIC = b"PK"
OLE_CFBF_MAGIC = b"\xd0\xcf\x11\xe0"

EXCEL_WRITE_HINT = "Fermez le fichier Excel/Calc, enregistrez vos modifications, puis réessayez."


class ExcelWriteError(Exception):
    """Fichier Excel inaccessible (ouvert ailleurs, verrouillé, format invalide…)."""


def thin_border() -> Border:
    side = Side(style="thin", color="B7E4C7")
    return Border(left=side, right=side, top=side, bottom=side)


def format_session_duration(seconds: float) -> str:
    """Affiche h/min/s sans tronquer les sessions courtes (ex. 5s → « 5s », pas « 0h 00min »)."""
    total = max(0, int(round(seconds)))
    hours = total // 3600
    minutes = (total % 3600) // 60
    secs = total % 60
    if hours > 0:
        return f"{hours}h {minutes:02}min {secs:02}s"
    if minutes > 0:
        return f"{minutes}min {secs:02}s"
    return f"{secs}s"


def sanitize_excel_text(value: str) -> str:
    """Empêche l'injection de formules Excel (=, +, -, @) dans les cellules texte.

    Retire d'abord tab/BOM/ZWSP/espaces en tête (contournements CSV/Excel),
    puis préfixe d'une apostrophe si le texte utile commence par un déclencheur.
    """
    if not value:
        return value
    cleaned = value.lstrip(_FORMULA_LEADING_STRIP)
    if cleaned and cleaned[0] in _FORMULA_PREFIXES:
        return f"'{cleaned}"
    return value


def _row_has_data(values: tuple[object, ...]) -> bool:
    return any(value is not None and str(value).strip() != "" for value in values)


def _session_row_values(ws: Worksheet, row_idx: int) -> tuple[object, ...]:
    return tuple(ws.cell(row=row_idx, column=col).value for col in range(1, 8))


def _last_row_with_data(ws: Worksheet, *, min_row: int = 2) -> int:
    """Dernière ligne avec données (max_row openpyxl peut être faux après édition)."""
    last = min_row - 1
    for row_idx, row in enumerate(ws.iter_rows(min_row=min_row, values_only=True), start=min_row):
        if _row_has_data(row):
            last = row_idx
    return last


def _next_session_row(ws: Worksheet) -> int:
    """Prochaine ligne vide après la dernière session (ne jamais écraser une ligne remplie)."""
    last = _last_row_with_data(ws, min_row=2)
    row_idx = max(last + 1, 2)
    while _row_has_data(_session_row_values(ws, row_idx)):
        row_idx += 1
    return row_idx


def _pure_path_for(excel_path: str) -> PurePath:
    """Choisit le parseur de chemin adapté (chemins Windows analysables sous Linux CI)."""
    if "\\" in excel_path or (len(excel_path) >= 2 and excel_path[1] == ":"):
        return PureWindowsPath(excel_path)
    return PurePosixPath(excel_path)


def _excel_lock_paths(excel_path: str) -> list[str]:
    pure = _pure_path_for(excel_path)
    filename = pure.name
    directory = pure.parent
    return [
        str(directory / f"~${filename}"),
        str(directory / f".~lock.{filename}#"),
    ]


def _assert_excel_ooxml(excel_path: str) -> None:
    """Refuse les faux .xlsx (ex. classeur OLE renommé / enregistré par Calc)."""
    if not os.path.exists(excel_path):
        return
    try:
        with open(excel_path, "rb") as handle:
            magic = handle.read(8)
    except OSError as exc:
        raise ExcelWriteError(
            f"Impossible de lire le fichier Excel :\n{exc}\n\n{EXCEL_WRITE_HINT}"
        ) from exc
    if magic.startswith(OOXML_ZIP_MAGIC):
        return
    if magic.startswith(OLE_CFBF_MAGIC):
        raise ExcelWriteError(
            "Le fichier n'est pas un vrai .xlsx (format Excel ancien / LibreOffice).\n\n"
            "Dans Calc : Fichier → Enregistrer sous → "
            "« Classeur Excel 2007-365 (.xlsx) ».\n"
            "Ou choisissez un nouvel emplacement dans Time Tracker."
        )
    raise ExcelWriteError(
        "Le fichier Excel est illisible ou corrompu.\n\n"
        "Choisissez un nouvel emplacement via « Changer l'emplacement… », "
        "ou réparez le fichier dans Calc."
    )


def _assert_excel_writable(excel_path: str) -> None:
    for lock_path in _excel_lock_paths(excel_path):
        if os.path.exists(lock_path):
            raise ExcelWriteError(
                f"Le fichier Excel semble ouvert dans une autre application.\n\n{EXCEL_WRITE_HINT}"
            )
    if not os.path.exists(excel_path):
        return
    _assert_excel_ooxml(excel_path)
    try:
        with open(excel_path, "r+b"):
            pass
    except PermissionError as exc:
        raise ExcelWriteError(
            f"Impossible d'écrire dans le fichier Excel.\n\n{EXCEL_WRITE_HINT}"
        ) from exc
    except OSError as exc:
        raise ExcelWriteError(
            f"Impossible d'accéder au fichier Excel :\n{exc}\n\n{EXCEL_WRITE_HINT}"
        ) from exc


def _save_workbook_atomic(wb: WorkbookType, excel_path: str) -> None:
    _assert_excel_writable(excel_path)
    parent_dir = os.path.dirname(excel_path)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)
    tmp_path = f"{excel_path}.tmp"
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, excel_path)
    except PermissionError as exc:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise ExcelWriteError(
            f"Impossible d'enregistrer le fichier Excel.\n\n{EXCEL_WRITE_HINT}"
        ) from exc
    except OSError as exc:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise ExcelWriteError(
            f"Impossible d'enregistrer le fichier Excel :\n{exc}\n\n{EXCEL_WRITE_HINT}"
        ) from exc


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
    for i, (header, width) in enumerate(zip(headers, col_widths, strict=True), 1):
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
    _save_workbook_atomic(wb, excel_path)


def log_session(
    project: str,
    start_dt: datetime,
    end_dt: datetime,
    notes: str,
    excel_path: str,
) -> int:
    """Ajoute une session en append. Retourne le numéro de ligne. Lève ExcelWriteError si échec."""
    ensure_excel(excel_path)
    _assert_excel_writable(excel_path)
    try:
        wb = load_workbook(excel_path)
    except Exception as exc:
        raise ExcelWriteError(
            "Impossible d'ouvrir le fichier Excel pour y écrire la session.\n\n"
            f"{exc}\n\n"
            "Vérifiez que le fichier est un vrai .xlsx et qu'il n'est pas ouvert ailleurs."
        ) from exc
    try:
        ws = wb["Sessions"]
    except KeyError as exc:
        raise ExcelWriteError(
            "Le fichier Excel ne contient pas la feuille « Sessions ».\n"
            "Choisissez un autre fichier ou recréez-en un via « Changer l'emplacement… »."
        ) from exc
    duration_h = (end_dt - start_dt).total_seconds() / 3600
    duration_min = (end_dt - start_dt).total_seconds() / 60
    row_idx = _next_session_row(ws)
    if _row_has_data(_session_row_values(ws, row_idx)):
        raise ExcelWriteError(
            f"La ligne {row_idx} de la feuille Sessions n'est pas vide.\n"
            "Aucune session n'a été écrasée."
        )
    fill = ALT_FILL if row_idx % 2 == 0 else WHITE_FILL
    values: list[str | float] = [
        start_dt.strftime("%d/%m/%Y"),
        sanitize_excel_text(project),
        start_dt.strftime("%H:%M"),
        end_dt.strftime("%H:%M"),
        round(duration_h, 4),
        round(duration_min, 1),
        sanitize_excel_text(notes),
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
    _save_workbook_atomic(wb, excel_path)
    return row_idx


def rename_project_in_excel(old_name: str, new_name: str, excel_path: str) -> None:
    if not os.path.exists(excel_path):
        return
    _assert_excel_writable(excel_path)
    try:
        wb = load_workbook(excel_path)
    except Exception as exc:
        raise ExcelWriteError(
            "Impossible d'ouvrir le fichier Excel pour renommer le projet.\n\n"
            f"{exc}\n\n{EXCEL_WRITE_HINT}"
        ) from exc
    try:
        ws = wb["Sessions"]
    except KeyError as exc:
        raise ExcelWriteError("Le fichier Excel ne contient pas la feuille « Sessions ».") from exc
    safe_new = sanitize_excel_text(new_name)
    for row in ws.iter_rows(min_row=2):
        if row[1].value == old_name or row[1].value == sanitize_excel_text(old_name):
            row[1].value = safe_new
    try:
        _refresh_summary(wb)
    except KeyError as exc:
        raise ExcelWriteError(
            "Le fichier Excel ne contient pas la feuille « Résumé par projet »."
        ) from exc
    _save_workbook_atomic(wb, excel_path)


def _parse_duration_hours(value: object) -> float | None:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return None


def _refresh_summary(wb: WorkbookType) -> None:
    ws = wb["Sessions"]
    ws2 = wb["Résumé par projet"]
    totals: dict[str, dict[str, float]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1]:
            project = str(row[1])
            hours = _parse_duration_hours(row[4])
            if hours is None:
                continue
            entry = totals.get(project, {"h": 0.0, "n": 0.0})
            entry["h"] += hours
            entry["n"] += 1
            totals[project] = entry
    sorted_projects = sorted(totals.items())
    last_summary = _last_row_with_data(ws2, min_row=3)
    for i in range(3, max(last_summary, 2 + len(sorted_projects)) + 1):
        _clear_summary_row(ws2, i)
    for i, (project, data) in enumerate(sorted_projects, start=3):
        fill2 = ALT_FILL if i % 2 == 0 else WHITE_FILL
        safe_project = sanitize_excel_text(str(project))
        values: list[str | float] = [safe_project, round(data["h"], 2), int(data["n"])]
        for col, val in enumerate(values, 1):
            cell = ws2.cell(row=i, column=col, value=val)
            cell.font = Font(name="Calibri", size=10)
            cell.fill = fill2
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center" if col != 1 else "left")
            if col == 1 and isinstance(val, str):
                cell.data_type = "s"
        ws2.cell(row=i, column=2).number_format = "0.00"


def _parse_session_date(value: object) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    try:
        return datetime.strptime(text, "%d/%m/%Y").date()
    except ValueError:
        return None


def _duration_label_from_times(date_text: str, start_text: str, end_text: str) -> str | None:
    try:
        start = datetime.strptime(f"{date_text} {start_text}", "%d/%m/%Y %H:%M")
        end = datetime.strptime(f"{date_text} {end_text}", "%d/%m/%Y %H:%M")
    except ValueError:
        return None
    if end < start:
        end = end + timedelta(days=1)
    return format_session_duration((end - start).total_seconds())


def _session_row_from_values(row_index: int, row: tuple[object, ...]) -> SessionRow | None:
    if not row or row[0] is None:
        return None
    date_text = str(row[0]).strip()
    if not date_text:
        return None
    start_text = str(row[2] or "").strip()
    end_text = str(row[3] or "").strip()
    # Préférer début/fin (évite cellules durée Excel incohérentes).
    duration_label = _duration_label_from_times(date_text, start_text, end_text)
    if duration_label is None:
        duration_h = _parse_duration_hours(row[4] if len(row) > 4 else None)
        if duration_h is None:
            duration_label = "?"
        else:
            duration_label = format_session_duration(duration_h * 3600)
    notes_raw = row[6] if len(row) > 6 else ""
    return {
        "row_index": row_index,
        "date": date_text,
        "project": str(row[1] or "").strip(),
        "start": start_text,
        "end": end_text,
        "duration": duration_label,
        "notes": str(notes_raw or "").strip(),
    }


def list_sessions_in_range(
    excel_path: str,
    start_day: date,
    end_day: date,
) -> list[SessionRow]:
    """Sessions dont la date est dans ``[start_day, end_day]`` inclus."""
    if end_day < start_day:
        start_day, end_day = end_day, start_day
    if not os.path.exists(excel_path):
        return []
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
    except Exception:
        return []
    try:
        try:
            ws = wb["Sessions"]
        except KeyError:
            return []
        rows: list[SessionRow] = []
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            parsed = _parse_session_date(row[0] if row else None)
            if parsed is None or parsed < start_day or parsed > end_day:
                continue
            session = _session_row_from_values(row_idx, row)
            if session is not None:
                rows.append(session)
        return rows
    finally:
        wb.close()


def list_sessions_for_day(
    excel_path: str,
    day: datetime | date | None = None,
) -> list[SessionRow]:
    """Lit les sessions Excel dont la date calendaire correspond à ``day``."""
    if day is None:
        target = datetime.now().date()
    elif isinstance(day, datetime):
        target = day.date()
    else:
        target = day
    return list_sessions_in_range(excel_path, target, target)


def week_bounds(day: date | None = None) -> tuple[date, date]:
    target = day or datetime.now().date()
    start = target - timedelta(days=target.weekday())
    return start, start + timedelta(days=6)


def month_bounds(day: date | None = None) -> tuple[date, date]:
    target = day or datetime.now().date()
    start = target.replace(day=1)
    if start.month == 12:
        next_month = start.replace(year=start.year + 1, month=1, day=1)
    else:
        next_month = start.replace(month=start.month + 1, day=1)
    end = next_month - timedelta(days=1)
    return start, end


def parse_session_datetimes(session: SessionRow) -> tuple[datetime, datetime]:
    start = datetime.strptime(f"{session['date']} {session['start']}", "%d/%m/%Y %H:%M")
    end = datetime.strptime(f"{session['date']} {session['end']}", "%d/%m/%Y %H:%M")
    if end < start:
        end = end + timedelta(days=1)
    return start, end


def total_session_hours(sessions: list[SessionRow]) -> float:
    total = 0.0
    for session in sessions:
        try:
            start, end = parse_session_datetimes(session)
            total += (end - start).total_seconds() / 3600.0
        except ValueError:
            continue
    return total


def hours_by_project(sessions: list[SessionRow]) -> list[tuple[str, float]]:
    """Totaux horaires par projet, triés décroissant puis nom."""
    totals: dict[str, float] = {}
    for session in sessions:
        try:
            start, end = parse_session_datetimes(session)
            hours = (end - start).total_seconds() / 3600.0
        except ValueError:
            continue
        name = session["project"] or "?"
        totals[name] = totals.get(name, 0.0) + hours
    return sorted(totals.items(), key=lambda item: (-item[1], item[0].casefold()))


def _apply_session_values(
    ws: Worksheet,
    row_idx: int,
    *,
    project: str,
    start_dt: datetime,
    end_dt: datetime,
    notes: str,
) -> None:
    duration_h = (end_dt - start_dt).total_seconds() / 3600
    duration_min = (end_dt - start_dt).total_seconds() / 60
    fill = ALT_FILL if row_idx % 2 == 0 else WHITE_FILL
    values: list[str | float] = [
        start_dt.strftime("%d/%m/%Y"),
        sanitize_excel_text(project),
        start_dt.strftime("%H:%M"),
        end_dt.strftime("%H:%M"),
        round(duration_h, 4),
        round(duration_min, 1),
        sanitize_excel_text(notes),
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


def update_session_row(
    excel_path: str,
    row_index: int,
    *,
    project: str,
    start_dt: datetime,
    end_dt: datetime,
    notes: str,
) -> None:
    if row_index < 2:
        raise ExcelWriteError("Index de ligne session invalide.")
    if end_dt <= start_dt:
        raise ExcelWriteError("L'heure de fin doit être après l'heure de début.")
    if not os.path.exists(excel_path):
        raise ExcelWriteError("Fichier Excel introuvable.")
    _assert_excel_writable(excel_path)
    try:
        wb = load_workbook(excel_path)
    except Exception as exc:
        raise ExcelWriteError(
            "Impossible d'ouvrir le fichier Excel pour modifier la session.\n\n"
            f"{exc}\n\n{EXCEL_WRITE_HINT}"
        ) from exc
    try:
        ws = wb["Sessions"]
    except KeyError as exc:
        raise ExcelWriteError("Le fichier Excel ne contient pas la feuille « Sessions ».") from exc
    if not _row_has_data(_session_row_values(ws, row_index)):
        raise ExcelWriteError(f"Aucune session à la ligne {row_index}.")
    _apply_session_values(
        ws,
        row_index,
        project=project,
        start_dt=start_dt,
        end_dt=end_dt,
        notes=notes,
    )
    try:
        _refresh_summary(wb)
    except KeyError as exc:
        raise ExcelWriteError(
            "Le fichier Excel ne contient pas la feuille « Résumé par projet »."
        ) from exc
    _save_workbook_atomic(wb, excel_path)


def delete_session_row(excel_path: str, row_index: int) -> None:
    if row_index < 2:
        raise ExcelWriteError("Index de ligne session invalide.")
    if not os.path.exists(excel_path):
        raise ExcelWriteError("Fichier Excel introuvable.")
    _assert_excel_writable(excel_path)
    try:
        wb = load_workbook(excel_path)
    except Exception as exc:
        raise ExcelWriteError(
            "Impossible d'ouvrir le fichier Excel pour supprimer la session.\n\n"
            f"{exc}\n\n{EXCEL_WRITE_HINT}"
        ) from exc
    try:
        ws = wb["Sessions"]
    except KeyError as exc:
        raise ExcelWriteError("Le fichier Excel ne contient pas la feuille « Sessions ».") from exc
    if not _row_has_data(_session_row_values(ws, row_index)):
        raise ExcelWriteError(f"Aucune session à la ligne {row_index}.")
    ws.delete_rows(row_index)
    try:
        _refresh_summary(wb)
    except KeyError as exc:
        raise ExcelWriteError(
            "Le fichier Excel ne contient pas la feuille « Résumé par projet »."
        ) from exc
    _save_workbook_atomic(wb, excel_path)


def open_excel_file(excel_path: str, *, parent: tk.Misc | None = None) -> bool:
    raw = excel_path.strip()
    if not raw:
        if parent is None:
            messagebox.showwarning("Chemin manquant", "Indiquez un chemin de fichier.")
        else:
            messagebox.showwarning(
                "Chemin manquant",
                "Indiquez un chemin de fichier.",
                parent=parent,
            )
        return False
    path = normalize_excel_path(raw)
    parent_dir = os.path.dirname(path)
    if parent_dir and not os.path.isdir(parent_dir):
        msg = f"Le dossier suivant n'existe pas :\n{parent_dir}"
        if parent is None:
            messagebox.showwarning("Dossier introuvable", msg)
        else:
            messagebox.showwarning("Dossier introuvable", msg, parent=parent)
        return False
    ensure_excel(path)
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except OSError as exc:
        err = f"Impossible d'ouvrir le fichier :\n{exc}"
        if parent is None:
            messagebox.showerror("Erreur", err)
        else:
            messagebox.showerror("Erreur", err, parent=parent)
        return False
    return True
