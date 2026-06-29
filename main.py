import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import time
import threading
import hashlib
from datetime import datetime, timedelta
from typing import Callable, Literal, TypedDict
import os
import sys
import json
import shutil
import subprocess
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SortMode = Literal["alpha_asc", "date_asc", "date_desc"]

class ProjectEntry(TypedDict):
    name: str
    created: str

class ProjectsConfig(TypedDict):
    projects: list[ProjectEntry]
    sort_mode: SortMode
    excel_file: str

DEFAULT_SORT: SortMode = "alpha_asc"
SORT_OPTIONS: list[tuple[str, SortMode]] = [
    ("A → Z", "alpha_asc"),
    ("Date ↑", "date_asc"),
    ("Date ↓", "date_desc"),
]
SORT_LABELS: dict[SortMode, str] = {mode: label for label, mode in SORT_OPTIONS}
APP_VERSION = "1.1.2"

INSTALL_DIR   = os.path.join(os.environ.get("LOCALAPPDATA", ""), "TimeTracker")
INSTALL_EXE   = os.path.join(INSTALL_DIR, "TimeTracker.exe")
REGISTRY_KEY  = r"Software\Microsoft\Windows\CurrentVersion\Run"
REGISTRY_NAME = "TimeTracker"
APP_NAME      = "Time Tracker"

def is_frozen() -> bool:
    return getattr(sys, "frozen", False)

def get_app_data_dir() -> str:
    if is_frozen():
        os.makedirs(INSTALL_DIR, exist_ok=True)
        return INSTALL_DIR
    return os.path.dirname(os.path.abspath(__file__))

def get_running_exe() -> str:
    if is_frozen():
        return os.path.abspath(sys.executable)
    return os.path.abspath(__file__)

def paths_equal(left: str, right: str) -> bool:
    return os.path.normcase(os.path.normpath(left)) == os.path.normcase(os.path.normpath(right))

def _register_startup(exe_path: str) -> None:
    if os.name != "nt":
        return
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, REGISTRY_NAME, 0, winreg.REG_SZ, f'"{exe_path}"')

def _create_shortcut(lnk_path: str, target: str, work_dir: str) -> None:
    if os.name != "nt":
        return
    os.makedirs(os.path.dirname(lnk_path), exist_ok=True)
    ps_lnk    = lnk_path.replace("'", "''")
    ps_target = target.replace("'", "''")
    ps_work   = work_dir.replace("'", "''")
    ps_name   = APP_NAME.replace("'", "''")
    script = (
        f"$s = (New-Object -ComObject WScript.Shell).CreateShortcut('{ps_lnk}'); "
        f"$s.TargetPath = '{ps_target}'; "
        f"$s.WorkingDirectory = '{ps_work}'; "
        f"$s.Description = '{ps_name}'; "
        f"$s.Save()"
    )
    creationflags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
    subprocess.run(
        ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", script],
        check=False,
        creationflags=creationflags,
    )

def _migrate_data(from_dir: str, to_dir: str) -> None:
    for filename in ("heures_projet.xlsx", "projects.json"):
        src = os.path.join(from_dir, filename)
        dst = os.path.join(to_dir, filename)
        if os.path.exists(src) and not os.path.exists(dst):
            shutil.copy2(src, dst)

def _file_hash(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(8192), b""):
            digest.update(chunk)
    return digest.hexdigest()

def ensure_installed() -> None:
    """Installe ou met a jour l'app dans AppData puis relance depuis la copie installee."""
    if not is_frozen() or os.name != "nt":
        return

    os.makedirs(INSTALL_DIR, exist_ok=True)
    running_exe = get_running_exe()

    if paths_equal(running_exe, INSTALL_EXE):
        _register_startup(INSTALL_EXE)
        return

    first_install = not os.path.exists(INSTALL_EXE)
    if first_install:
        # Uniquement a la premiere install : ne jamais ecraser projects.json / Excel existants.
        _migrate_data(os.path.dirname(running_exe), INSTALL_DIR)
    needs_update = (
        not first_install
        and _file_hash(running_exe) != _file_hash(INSTALL_EXE)
    )
    shutil.copy2(running_exe, INSTALL_EXE)
    _register_startup(INSTALL_EXE)

    desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop", f"{APP_NAME}.lnk")
    start_menu = os.path.join(
        os.environ.get("APPDATA", ""),
        "Microsoft", "Windows", "Start Menu", "Programs",
        f"{APP_NAME}.lnk",
    )
    _create_shortcut(desktop, INSTALL_EXE, INSTALL_DIR)
    _create_shortcut(start_menu, INSTALL_EXE, INSTALL_DIR)

    if first_install or needs_update:
        root = tk.Tk()
        root.withdraw()
        if first_install:
            message = (
                f"Installation terminee.\n\n"
                f"Dossier : {INSTALL_DIR}\n"
                f"Raccourci Bureau + Menu Demarrer crees.\n"
                f"L'app se lancera a chaque demarrage de Windows."
            )
        else:
            message = (
                f"Mise a jour terminee (v{APP_VERSION}).\n\n"
                f"Dossier : {INSTALL_DIR}\n"
                f"L'application va redemarrer avec la nouvelle version."
            )
        messagebox.showinfo(APP_NAME, message, parent=root)
        root.destroy()

    creationflags = subprocess.DETACHED_PROCESS if hasattr(subprocess, "DETACHED_PROCESS") else 0
    subprocess.Popen([INSTALL_EXE], close_fds=True, creationflags=creationflags)
    sys.exit(0)

APP_DIR     = get_app_data_dir()
CONFIG_FILE = os.path.join(APP_DIR, "projects.json")
DEFAULT_EXCEL_NAME = "heures_projet.xlsx"

def default_excel_path() -> str:
    return os.path.join(APP_DIR, DEFAULT_EXCEL_NAME)

def normalize_excel_path(path: str) -> str:
    cleaned = os.path.abspath(os.path.expanduser(path.strip()))
    if not cleaned.lower().endswith(".xlsx"):
        cleaned = f"{cleaned}.xlsx"
    return cleaned

# ─── CONFIG ────────────────────────────────────────────────────────────────────

def _default_projects() -> list[ProjectEntry]:
    now = datetime.now().isoformat(timespec="seconds")
    return [
        {"name": "Projet 1", "created": now},
        {"name": "Projet 2", "created": now},
        {"name": "Projet 3", "created": now},
    ]

def _normalize_sort_mode(value: str) -> SortMode:
    if value in ("alpha_asc", "date_asc", "date_desc"):
        return value  # type: ignore[return-value]
    return DEFAULT_SORT

def _migrate_legacy_projects(raw: list[str]) -> list[ProjectEntry]:
    base = datetime.now().replace(microsecond=0)
    return [
        {"name": name, "created": (base + timedelta(seconds=index)).isoformat(timespec="seconds")}
        for index, name in enumerate(raw)
    ]

def load_config() -> ProjectsConfig:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as handle:
            raw = json.load(handle)
        if isinstance(raw, list):
            return {
                "projects": _migrate_legacy_projects(raw),
                "sort_mode": DEFAULT_SORT,
                "excel_file": default_excel_path(),
            }
        projects = raw.get("projects", _default_projects())
        normalized: list[ProjectEntry] = []
        for item in projects:
            if isinstance(item, str):
                normalized.append({"name": item, "created": datetime.now().isoformat(timespec="seconds")})
            else:
                normalized.append({
                    "name": str(item["name"]),
                    "created": str(item.get("created", datetime.now().isoformat(timespec="seconds"))),
                })
        excel_raw = raw.get("excel_file")
        excel_file = normalize_excel_path(str(excel_raw)) if excel_raw else default_excel_path()
        return {
            "projects": normalized,
            "sort_mode": _normalize_sort_mode(str(raw.get("sort_mode", DEFAULT_SORT))),
            "excel_file": excel_file,
        }
    return {
        "projects": _default_projects(),
        "sort_mode": DEFAULT_SORT,
        "excel_file": default_excel_path(),
    }

def save_config(config: ProjectsConfig) -> None:
    payload: ProjectsConfig = {
        "projects": config["projects"],
        "sort_mode": config["sort_mode"],
        "excel_file": normalize_excel_path(config["excel_file"]),
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)

def sort_projects(projects: list[ProjectEntry], mode: SortMode) -> list[ProjectEntry]:
    if mode == "alpha_asc":
        return sorted(projects, key=lambda entry: entry["name"].casefold())
    if mode == "date_asc":
        return sorted(projects, key=lambda entry: entry["created"])
    return sorted(projects, key=lambda entry: entry["created"], reverse=True)

def project_names(projects: list[ProjectEntry]) -> list[str]:
    return [entry["name"] for entry in projects]

class SearchableDropdown(tk.Frame):
    """Select avec recherche integree dans le panneau deroulant."""

    POPUP_LIST_HEIGHT = 8

    def __init__(
        self,
        parent: tk.Misc,
        textvariable: tk.StringVar,
        *,
        bg: str,
        field_bg: str,
        border: str,
        muted: str,
        accent: str,
        highlight: str,
        text: str,
        empty_label: str = "Aucun projet",
        placeholder: str = "Choisir un projet…",
    ) -> None:
        super().__init__(parent, bg=bg)
        self.variable = textvariable
        self._bg = bg
        self._field_bg = field_bg
        self._border = border
        self._muted = muted
        self._accent = accent
        self._highlight = highlight
        self._text = text
        self._empty_label = empty_label
        self._placeholder = placeholder
        self._items: list[str] = []
        self._popup: tk.Toplevel | None = None
        self._listbox: tk.Listbox | None = None
        self._search_var = tk.StringVar()
        self._search_trace_id: str | None = None
        self._click_bind_id: str | None = None
        self._open = False
        self._updating_list = False
        self._outside_click_armed = False

        shell = tk.Frame(self, bg=border, padx=1, pady=1)
        shell.pack(fill="x")
        self.trigger = tk.Frame(shell, bg=field_bg, cursor="hand2")
        self.trigger.pack(fill="x")

        self.value_label = tk.Label(
            self.trigger, text="", bg=field_bg, fg=text, anchor="w",
            font=("Segoe UI", 10), padx=10, pady=8,
        )
        self.value_label.pack(side="left", fill="x", expand=True)
        caret = tk.Label(
            self.trigger, text="▼", bg=field_bg, fg=muted,
            font=("Segoe UI", 8),
        )
        caret.pack(side="right", padx=(0, 10))

        for widget in (self.trigger, self.value_label, caret):
            widget.bind("<Button-1>", self._toggle)

        self.variable.trace_add("write", lambda *_: self._update_display())
        self._update_display()

    def _update_display(self) -> None:
        value = self.variable.get().strip()
        if value:
            self.value_label.config(text=value, fg=self._text)
        else:
            self.value_label.config(text=self._placeholder, fg=self._muted)

    def set_options(self, options: list[str]) -> None:
        self._items = list(options)
        current = self.variable.get()
        if current and current not in self._items:
            if self._items:
                self.variable.set(self._items[0])
            else:
                self.variable.set("")
        elif not current and self._items:
            self.variable.set(self._items[0])
        if self._open:
            self._refresh_list()

    def close_if_open(self) -> None:
        if self._open:
            self._close()

    def _toggle(self, _event: tk.Event | None = None) -> None:
        if self._open:
            self._close()
        else:
            self._open_popup()

    def _open_popup(self) -> None:
        if not self._items or self._open:
            return

        self._close()
        self._popup = tk.Toplevel(self)
        self._popup.overrideredirect(True)
        self._popup.configure(bg=self._border)
        self._popup.attributes("-topmost", True)

        width = max(self.winfo_width(), 260)
        popup_height = 44 + self.POPUP_LIST_HEIGHT * 24 + 8
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 2
        self._popup.geometry(f"{width}x{popup_height}+{x}+{y}")

        container = tk.Frame(self._popup, bg=self._bg, padx=1, pady=1)
        container.pack(fill="both", expand=True)

        search_wrap = tk.Frame(container, bg=self._field_bg)
        search_wrap.pack(fill="x")
        self._search_var.set("")
        search_entry = tk.Entry(
            search_wrap, textvariable=self._search_var,
            font=("Segoe UI", 10), bg=self._field_bg, fg=self._text,
            relief="flat", bd=0, highlightthickness=1,
            highlightbackground=self._border, highlightcolor=self._highlight,
        )
        search_entry.pack(fill="x", padx=8, pady=8, ipady=3)
        search_entry.bind("<Escape>", lambda _event: self._close())
        search_entry.bind("<Return>", self._select_first)
        search_entry.bind("<Down>", lambda _event: self._focus_list())

        list_shell = tk.Frame(container, bg=self._border, padx=1)
        list_shell.pack(fill="both", expand=True, pady=(0, 1))
        list_frame = tk.Frame(list_shell, bg=self._bg)
        list_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self._listbox = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set,
            font=("Segoe UI", 10), bg=self._bg, fg=self._text,
            selectbackground=self._highlight, selectforeground="white",
            activestyle="none", relief="flat", bd=0,
            highlightthickness=0, height=self.POPUP_LIST_HEIGHT,
        )
        scrollbar.config(command=self._listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self._listbox.pack(side="left", fill="both", expand=True)

        self._listbox.bind("<Double-Button-1>", self._on_list_select)
        self._listbox.bind("<Return>", self._on_list_select)
        self._listbox.bind("<Escape>", lambda _event: self._close())
        self._search_trace_id = self._search_var.trace_add(
            "write", lambda *_: self._refresh_list(),
        )

        self._open = True
        self._outside_click_armed = False
        self._refresh_list()
        search_entry.focus_set()

        root = self.winfo_toplevel()
        self.after(120, lambda: self._arm_outside_click(root))

    def _arm_outside_click(self, root: tk.Misc) -> None:
        if not self._open:
            return
        self._outside_click_armed = True
        self._click_bind_id = root.bind("<Button-1>", self._on_outside_click, add="+")
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.bind("<Escape>", lambda _event: self._close())

    def _filtered_options(self) -> list[str]:
        query = self._search_var.get().casefold().strip()
        if not query:
            return self._items
        return [name for name in self._items if query in name.casefold()]

    def _listbox_alive(self) -> bool:
        return self._listbox is not None and self._listbox.winfo_exists()

    def _refresh_list(self) -> None:
        if not self._open or not self._listbox_alive():
            return
        filtered = self._filtered_options()
        current = self.variable.get()
        self._updating_list = True
        self._listbox.delete(0, "end")
        if not filtered:
            self._listbox.insert("end", self._empty_label)
            self._listbox.config(state="disabled")
            self._updating_list = False
            return
        self._listbox.config(state="normal")
        for name in filtered:
            self._listbox.insert("end", name)
        for index, name in enumerate(filtered):
            if name == current:
                self._listbox.selection_set(index)
                self._listbox.see(index)
                break
        self._updating_list = False

    def _focus_list(self) -> None:
        if not self._listbox_alive():
            return
        if self._listbox.size() > 0 and str(self._listbox.get(0)) != self._empty_label:
            self._listbox.focus_set()
            if not self._listbox.curselection():
                self._listbox.selection_set(0)

    def _select_first(self, _event: tk.Event | None = None) -> None:
        filtered = self._filtered_options()
        if filtered:
            self.variable.set(filtered[0])
        self._close()

    def _on_list_select(self, _event: tk.Event | None = None) -> None:
        if self._updating_list or not self._listbox_alive():
            return
        if self._listbox.cget("state") == "disabled":
            return
        selection = self._listbox.curselection()
        if not selection:
            return
        value = self._listbox.get(selection[0])
        if value == self._empty_label:
            return
        self.variable.set(value)
        self._close()

    def _on_outside_click(self, event: tk.Event) -> None:
        if not self._open or not self._outside_click_armed or self._popup is None:
            return
        if self._contains(self.trigger, event.x_root, event.y_root):
            return
        if self._contains(self._popup, event.x_root, event.y_root):
            return
        self._close()

    def _contains(self, widget: tk.Misc, x_root: int, y_root: int) -> bool:
        if not widget.winfo_exists():
            return False
        x = widget.winfo_rootx()
        y = widget.winfo_rooty()
        return x <= x_root <= x + widget.winfo_width() and y <= y_root <= y + widget.winfo_height()

    def _close(self) -> None:
        if self._search_trace_id is not None:
            self._search_var.trace_remove("write", self._search_trace_id)
            self._search_trace_id = None
        if self._click_bind_id is not None:
            self.winfo_toplevel().unbind("<Button-1>", self._click_bind_id)
            self._click_bind_id = None
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        self._listbox = None
        self._open = False
        self._outside_click_armed = False

# ─── EXCEL ─────────────────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", fgColor="2D6A4F")
SUBHEAD_FILL = PatternFill("solid", fgColor="40916C")
ALT_FILL     = PatternFill("solid", fgColor="F0FFF4")
WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")

def thin_border():
    s = Side(style="thin", color="B7E4C7")
    return Border(left=s, right=s, top=s, bottom=s)

def _row_has_data(values: tuple[object, ...]) -> bool:
    return any(value is not None and str(value).strip() != "" for value in values)

def _last_row_with_data(ws, *, min_row: int = 2) -> int:
    """Dernière ligne contenant des données (openpyxl.max_row peut être incorrect après édition manuelle)."""
    last = min_row - 1
    for row_idx, row in enumerate(ws.iter_rows(min_row=min_row, values_only=True), start=min_row):
        if _row_has_data(row):
            last = row_idx
    return last

def _next_session_row(ws) -> int:
    last = _last_row_with_data(ws, min_row=2)
    return max(last + 1, 2)

def _clear_summary_row(ws2, row_idx: int) -> None:
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
    ws.title = "Sessions"
    headers    = ["Date", "Projet", "Début", "Fin", "Durée (h)", "Durée (min)", "Notes"]
    col_widths = [14, 22, 10, 10, 12, 14, 30]
    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = thin_border()
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("Résumé par projet")
    ws2["A1"] = "Résumé par Projet"
    ws2["A1"].font      = Font(bold=True, color="FFFFFF", name="Calibri", size=14)
    ws2["A1"].fill      = HEADER_FILL
    ws2["A1"].alignment = Alignment(horizontal="center")
    ws2.merge_cells("A1:C1")
    ws2.row_dimensions[1].height = 28
    for i, h in enumerate(["Projet", "Total heures", "Nb sessions"], 1):
        cell = ws2.cell(row=2, column=i, value=h)
        cell.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.fill      = SUBHEAD_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border    = thin_border()
    ws2.column_dimensions["A"].width = 24
    ws2.column_dimensions["B"].width = 16
    ws2.column_dimensions["C"].width = 14
    ws2.freeze_panes = "A3"
    wb.save(excel_path)

def log_session(project: str, start_dt: datetime, end_dt: datetime, notes: str, excel_path: str) -> None:
    ensure_excel(excel_path)
    wb = load_workbook(excel_path)
    ws = wb["Sessions"]
    duration_h   = (end_dt - start_dt).total_seconds() / 3600
    duration_min = (end_dt - start_dt).total_seconds() / 60
    nr   = _next_session_row(ws)
    fill = ALT_FILL if nr % 2 == 0 else WHITE_FILL
    values = [
        start_dt.strftime("%d/%m/%Y"), project,
        start_dt.strftime("%H:%M"), end_dt.strftime("%H:%M"),
        round(duration_h, 4), round(duration_min, 1), notes,
    ]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=nr, column=col, value=val)
        cell.font      = Font(name="Calibri", size=10)
        cell.fill      = fill
        cell.border    = thin_border()
        cell.alignment = Alignment(horizontal="center" if col != 7 else "left", vertical="center")
    ws.cell(row=nr, column=5).number_format = "0.00"
    ws.cell(row=nr, column=6).number_format = "0.0"
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

def _refresh_summary(wb):
    ws  = wb["Sessions"]
    ws2 = wb["Résumé par projet"]
    totals: dict[str, dict[str, float | int]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1]:
            p = str(row[1])
            h = float(row[4] or 0)
            totals[p] = totals.get(p, {"h": 0.0, "n": 0})
            totals[p]["h"] += h
            totals[p]["n"] += 1
    sorted_projects = sorted(totals.items())
    auto_end = 2 + len(sorted_projects)
    for i in range(3, auto_end + 1):
        _clear_summary_row(ws2, i)
    # Supprimer d'anciennes lignes auto en doublon sans toucher aux lignes manuelles.
    for i in range(auto_end + 1, _last_row_with_data(ws2, min_row=3) + 1):
        name = ws2.cell(row=i, column=1).value
        if name is not None and str(name) in totals:
            _clear_summary_row(ws2, i)
    for i, (p, data) in enumerate(sorted_projects, start=3):
        fill2 = ALT_FILL if i % 2 == 0 else WHITE_FILL
        for col, val in enumerate([p, round(data["h"], 2), data["n"]], 1):
            cell = ws2.cell(row=i, column=col, value=val)
            cell.font      = Font(name="Calibri", size=10)
            cell.fill      = fill2
            cell.border    = thin_border()
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

# ─── EXCEL FILE DIALOG ─────────────────────────────────────────────────────────

class ExcelFileDialog(tk.Toplevel):
    """Fenêtre modale pour choisir l'emplacement et le nom du fichier Excel."""

    def __init__(self, parent: tk.Misc, excel_file: str, on_save: Callable[[str], None]) -> None:
        super().__init__(parent)
        self._on_save = on_save
        self._path_var = tk.StringVar(value=excel_file)
        self.title("Fichier Excel")
        self.resizable(True, True)
        self.minsize(360, 200)
        self.grab_set()
        self.attributes("-topmost", True)
        self.configure(bg="#FFFFFF")
        self._build()
        self.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        self.geometry(f"500x220+{px - 250}+{py}")

    def _build(self) -> None:
        BG = "#FFFFFF"
        ACCENT = "#40916C"
        GREEN = "#52B788"
        MUTED = "#888888"
        BORDER = "#E0E0E0"

        hdr = tk.Frame(self, bg=ACCENT, height=36)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📁  Fichier Excel", bg=ACCENT, fg="white",
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=12, pady=6)

        body = tk.Frame(self, bg=BG, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        tk.Label(
            body,
            text="Emplacement et nom du fichier d'export :",
            bg=BG, fg=MUTED, font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", pady=(0, 6))

        path_row = tk.Frame(body, bg=BG)
        path_row.pack(fill="x", pady=(0, 8))
        self._path_entry = tk.Entry(
            path_row, textvariable=self._path_var,
            font=("Segoe UI", 9), bg="#F5F5F5", fg="#222222",
            relief="flat", bd=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=GREEN,
        )
        self._path_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))
        tk.Button(
            path_row, text="Ouvrir", bg="#F0F0F0", fg="#444444", bd=0,
            font=("Segoe UI", 9), cursor="hand2",
            padx=10, pady=6, command=self._open,
        ).pack(side="left", padx=(0, 6))
        tk.Button(
            path_row, text="Parcourir…", bg=GREEN, fg="white", bd=0,
            font=("Segoe UI", 9, "bold"), cursor="hand2",
            padx=10, pady=6, activebackground=ACCENT,
            activeforeground="white", command=self._browse,
        ).pack(side="left")

        tk.Label(
            body,
            text="Les nouvelles sessions seront enregistrées à cet emplacement.\n"
                 "L'ancien fichier n'est pas déplacé automatiquement.",
            bg=BG, fg=MUTED, font=("Segoe UI", 8), justify="left",
        ).pack(anchor="w", pady=(0, 10))

        btn_row = tk.Frame(body, bg=BG)
        btn_row.pack(fill="x")

        tk.Button(
            btn_row, text="Par défaut", bg="#F0F0F0", fg="#444444", bd=0,
            font=("Segoe UI", 9), cursor="hand2", padx=10, pady=6,
            command=self._use_default,
        ).pack(side="left")

        tk.Button(
            btn_row, text="Annuler", bg="#F0F0F0", fg="#444444", bd=0,
            font=("Segoe UI", 9), cursor="hand2", padx=10, pady=6,
            command=self.destroy,
        ).pack(side="right", padx=(6, 0))

        tk.Button(
            btn_row, text="Enregistrer", bg=ACCENT, fg="white", bd=0,
            font=("Segoe UI", 9, "bold"), cursor="hand2", padx=12, pady=6,
            activebackground=GREEN, activeforeground="white",
            command=self._save,
        ).pack(side="right")

    def _browse(self) -> None:
        current = self._path_var.get().strip() or default_excel_path()
        chosen = filedialog.asksaveasfilename(
            parent=self,
            title="Choisir l'emplacement du fichier Excel",
            defaultextension=".xlsx",
            filetypes=[("Fichier Excel", "*.xlsx"), ("Tous les fichiers", "*.*")],
            initialfile=os.path.basename(current),
            initialdir=os.path.dirname(current) or APP_DIR,
        )
        if chosen:
            self._path_var.set(normalize_excel_path(chosen))

    def _use_default(self) -> None:
        self._path_var.set(default_excel_path())

    def _open(self) -> None:
        open_excel_file(self._path_var.get(), parent=self)

    def _save(self) -> None:
        raw = self._path_var.get().strip()
        if not raw:
            messagebox.showwarning("Chemin manquant", "Indiquez un chemin de fichier.", parent=self)
            return
        excel_path = normalize_excel_path(raw)
        parent_dir = os.path.dirname(excel_path)
        if parent_dir and not os.path.isdir(parent_dir):
            if not messagebox.askyesno(
                "Dossier inexistant",
                f"Le dossier suivant n'existe pas :\n{parent_dir}\n\nLe créer ?",
                parent=self,
            ):
                return
            try:
                os.makedirs(parent_dir, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Erreur", f"Impossible de créer le dossier :\n{exc}", parent=self)
                return
        self._on_save(excel_path)
        self.destroy()

# ─── MANAGE PROJECTS DIALOG ────────────────────────────────────────────────────

class ManageProjectsDialog(tk.Toplevel):
    """Fenêtre modale pour gérer (modifier / supprimer) les projets."""

    def __init__(
        self,
        parent: tk.Misc,
        projects: list[ProjectEntry],
        sort_mode: SortMode,
        excel_file: str,
        on_close,
    ):
        super().__init__(parent)
        self.projects: list[ProjectEntry] = list(projects)
        self.excel_file = excel_file
        self.sort_mode = tk.StringVar(value=SORT_LABELS[sort_mode])
        names = project_names(sort_projects(self.projects, sort_mode))
        self.selected_name = tk.StringVar(value=names[0] if names else "")
        self.on_close = on_close
        self.title("Gérer les projets")
        self.resizable(True, True)
        self.minsize(320, 320)
        self.grab_set()
        self.attributes("-topmost", True)
        self.configure(bg="#FFFFFF")
        self._build()
        self._refresh_picker()
        self.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        self.geometry(f"320x260+{px - 260}+{py}")

    def _build(self):
        BG     = "#FFFFFF"
        ACCENT = "#40916C"
        GREEN  = "#52B788"
        RED    = "#E63946"
        MUTED  = "#888888"
        BORDER = "#E0E0E0"

        hdr = tk.Frame(self, bg=ACCENT, height=36)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="✏  Gérer les projets", bg=ACCENT, fg="white",
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=12, pady=6)

        body = tk.Frame(self, bg=BG, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        top_row = tk.Frame(body, bg=BG)
        top_row.pack(fill="x", pady=(0, 6))
        tk.Label(top_row, text="Projet", bg=BG, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(side="left")

        sort_combo = ttk.Combobox(
            top_row, textvariable=self.sort_mode, state="readonly",
            values=[label for label, _mode in SORT_OPTIONS],
            font=("Segoe UI", 8), width=8,
        )
        sort_combo.pack(side="right", padx=(8, 0))
        sort_combo.bind("<<ComboboxSelected>>", self._on_sort_selected)

        self.project_picker = SearchableDropdown(
            body, self.selected_name,
            bg=BG, field_bg="#F5F5F5", border=BORDER,
            muted=MUTED, accent=GREEN, highlight=GREEN, text="#222222",
            placeholder="Choisir un projet…",
        )
        self.project_picker.pack(fill="x", pady=(0, 12))
        self.bind("<Configure>", self._on_dialog_configure)

        btn_frame = tk.Frame(body, bg=BG)
        btn_frame.pack(fill="x", pady=(0, 0))

        def btn(parent, text, color, cmd):
            button = tk.Button(parent, text=text, bg=color, fg="white", bd=0,
                               font=("Segoe UI", 9, "bold"), cursor="hand2",
                               padx=10, pady=6, activebackground=ACCENT,
                               activeforeground="white", command=cmd)
            button.pack(side="left", padx=(0, 6))
            return button

        btn(btn_frame, "✏  Renommer", ACCENT, self._rename)
        btn(btn_frame, "🗑  Supprimer", RED, self._delete)

        tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=(12, 8))

        tk.Label(body, text="Ajouter un projet", bg=BG, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 4))
        add_row = tk.Frame(body, bg=BG)
        add_row.pack(fill="x")
        self.new_var = tk.StringVar()
        entry = tk.Entry(add_row, textvariable=self.new_var,
                         font=("Segoe UI", 10), bg="#F5F5F5", fg="#222222",
                         relief="flat", bd=0, highlightthickness=1,
                         highlightbackground=BORDER, highlightcolor=GREEN)
        entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))
        entry.bind("<Return>", lambda _event: self._add())
        tk.Button(add_row, text="＋ Ajouter", bg=GREEN, fg="white", bd=0,
                  font=("Segoe UI", 9, "bold"), cursor="hand2",
                  padx=10, pady=6, activebackground=ACCENT,
                  activeforeground="white", command=self._add).pack(side="left")

        tk.Button(body, text="Fermer", bg="#F0F0F0", fg="#444444", bd=0,
                  font=("Segoe UI", 9), cursor="hand2", padx=10, pady=6,
                  command=self._close).pack(anchor="e", pady=(12, 0))

    def _current_sort_mode(self) -> SortMode:
        label = self.sort_mode.get()
        for option_label, mode in SORT_OPTIONS:
            if option_label == label:
                return mode
        return DEFAULT_SORT

    def _on_dialog_configure(self, event: tk.Event) -> None:
        if event.widget is self:
            self.project_picker.close_if_open()

    def _on_sort_selected(self, _event=None) -> None:
        self._refresh_picker()

    def _sorted_names(self) -> list[str]:
        ordered = sort_projects(self.projects, self._current_sort_mode())
        return project_names(ordered)

    def _refresh_picker(self) -> None:
        self.project_picker.set_options(self._sorted_names())

    def _selected_name(self) -> str | None:
        name = self.selected_name.get().strip()
        if name and name in project_names(self.projects):
            return name
        return None

    def _require_selection(self) -> str | None:
        name = self._selected_name()
        if not name:
            messagebox.showinfo("Sélection", "Sélectionnez d'abord un projet.", parent=self)
            return None
        return name

    def _find_project(self, name: str) -> ProjectEntry | None:
        for entry in self.projects:
            if entry["name"] == name:
                return entry
        return None

    def _rename(self):
        old_name = self._require_selection()
        if old_name is None:
            return
        old_entry = self._find_project(old_name)
        if old_entry is None:
            return
        new = simpledialog.askstring(
            "Renommer", f"Nouveau nom pour « {old_name} » :",
            initialvalue=old_name, parent=self,
        )
        if not new or not new.strip() or new.strip() == old_name:
            return
        new_name = new.strip()
        if any(entry["name"] == new_name for entry in self.projects):
            messagebox.showwarning("Doublon", f"« {new_name} » existe déjà.", parent=self)
            return
        old_entry["name"] = new_name
        self.selected_name.set(new_name)
        self._refresh_picker()
        rename_project_in_excel(old_name, new_name, self.excel_file)

    def _delete(self):
        name = self._require_selection()
        if name is None:
            return
        if not messagebox.askyesno(
            "Confirmer",
            f"Supprimer « {name} » ?\n\nLes sessions Excel seront conservées.",
            parent=self,
        ):
            return
        self.projects = [entry for entry in self.projects if entry["name"] != name]
        self._refresh_picker()

    def _add(self):
        name = self.new_var.get().strip()
        if not name:
            return
        if any(entry["name"] == name for entry in self.projects):
            messagebox.showwarning("Doublon", f"« {name} » existe déjà.", parent=self)
            return
        self.projects.append({
            "name": name,
            "created": datetime.now().isoformat(timespec="seconds"),
        })
        self.new_var.set("")
        self.selected_name.set(name)
        self._refresh_picker()

    def _close(self):
        self.destroy()

    def destroy(self):
        self.project_picker.close_if_open()
        save_config({
            "projects": self.projects,
            "sort_mode": self._current_sort_mode(),
            "excel_file": self.excel_file,
        })
        self.on_close(self.projects, self._current_sort_mode())
        self.grab_release()
        super().destroy()

# ─── MAIN APP ──────────────────────────────────────────────────────────────────

class TimeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("⏱ Time Tracker")
        self.root.minsize(240, 370)
        self.root.geometry("240x370")
        self.root.resizable(True, True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.98)

        config = load_config()
        self.projects: list[ProjectEntry] = config["projects"]
        self.excel_file: str = config["excel_file"]
        self.sort_mode = tk.StringVar(value=SORT_LABELS[config["sort_mode"]])
        names = project_names(sort_projects(self.projects, config["sort_mode"]))
        self.selected_project = tk.StringVar(value=names[0] if names else "")
        self.tracking         = False
        self.start_time       = None
        self.elapsed          = 0
        self._running         = False

        self._build_ui()
        self._refresh_project_list()
        self._position_window()

    def _position_window(self):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w  = self.root.winfo_width()
        h  = self.root.winfo_height()
        self.root.geometry(f"+{sw - w - 30}+{sh - h - 60}")

    def _build_ui(self):
        # ── Palette (light / white theme) ──
        BG     = "#FFFFFF"
        CARD   = "#FFFFFF"
        GREEN  = "#40916C"
        RED    = "#E63946"
        TEXT   = "#1A1A1A"
        MUTED  = "#666666"
        ACCENT = "#2D6A4F"
        SUBTLE = "#F5F5F5"
        BORDER = "#E0E0E0"
        TIMER_COLOR = "#1B4332"

        self._BG     = BG
        self._CARD   = CARD
        self._GREEN  = GREEN
        self._RED    = RED
        self._TEXT   = TEXT
        self._MUTED  = MUTED
        self._ACCENT = ACCENT
        self._SUBTLE = SUBTLE
        self._BORDER = BORDER
        self._TIMER  = TIMER_COLOR

        self.root.configure(bg=BG)

        # ── Title bar ──
        title_bar = tk.Frame(self.root, bg=ACCENT, height=38)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text="⏱  Time Tracker", bg=ACCENT, fg="white",
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=12, pady=7)

        for widget in [title_bar] + title_bar.winfo_children():
            if not isinstance(widget, tk.Button):
                widget.bind("<ButtonPress-1>", self._start_drag)
                widget.bind("<B1-Motion>",     self._drag)

        # ── Body ──
        body = tk.Frame(self.root, bg=BG, padx=16, pady=14)
        body.pack(fill="both", expand=True)

        # Project label + manage link
        top_row = tk.Frame(body, bg=BG)
        top_row.pack(fill="x", pady=(0, 6))
        tk.Label(top_row, text="PROJET", bg=BG, fg=MUTED,
                 font=("Segoe UI", 7, "bold")).pack(side="left")

        style = ttk.Style()
        style.theme_use("default")
        style.configure("W.TCombobox",
                        fieldbackground=SUBTLE, background=SUBTLE,
                        foreground=TEXT, selectbackground=GREEN,
                        selectforeground="white", arrowcolor=ACCENT,
                        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
        style.map("W.TCombobox", fieldbackground=[("readonly", SUBTLE)])

        sort_combo = ttk.Combobox(
            top_row, textvariable=self.sort_mode, state="readonly",
            values=[label for label, _mode in SORT_OPTIONS],
            font=("Segoe UI", 8), width=8, style="W.TCombobox",
        )
        sort_combo.pack(side="right", padx=(8, 0))
        sort_combo.bind("<<ComboboxSelected>>", self._on_sort_changed)

        tk.Button(top_row, text="Gérer ›", bg=BG, fg=GREEN, bd=0,
                  font=("Segoe UI", 7, "underline"), cursor="hand2",
                  activebackground=BG, activeforeground=ACCENT,
                  command=self._open_manage).pack(side="right")

        self.project_picker = SearchableDropdown(
            body, self.selected_project,
            bg=BG, field_bg=SUBTLE, border=BORDER,
            muted=MUTED, accent=ACCENT, highlight=GREEN, text=TEXT,
        )
        self.project_picker.pack(fill="x", pady=(0, 12))
        self.root.bind("<Configure>", self._on_root_configure)

        tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=(0, 12))

        self.timer_label = tk.Label(body, text="00:00:00", bg=BG, fg=TIMER_COLOR,
                                     font=("Segoe UI Mono", 32, "bold"))
        self.timer_label.pack()

        self.status_label = tk.Label(body, text="● En attente", bg=BG, fg=MUTED,
                                      font=("Segoe UI", 9))
        self.status_label.pack(pady=(2, 14))

        btn_row = tk.Frame(body, bg=BG)
        btn_row.pack(fill="x")

        self.start_btn = tk.Button(
            btn_row, text="▶  Start", bg=GREEN, fg="white",
            font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2",
            padx=14, pady=9, activebackground=ACCENT,
            activeforeground="white", command=self._start
        )
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.stop_btn = tk.Button(
            btn_row, text="■  Stop", bg=SUBTLE, fg=MUTED,
            font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2",
            padx=14, pady=9, activebackground=RED,
            activeforeground="white", command=self._stop, state="disabled"
        )
        self.stop_btn.pack(side="left", fill="x", expand=True)

        tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=(12, 10))

        excel_footer = tk.Frame(body, bg=BG)
        excel_footer.pack(fill="x", pady=(0, 6))

        def excel_link(parent: tk.Frame, text: str, command) -> None:
            tk.Button(
                parent, text=text, bg=BG, fg=MUTED,
                font=("Segoe UI", 8), bd=0, cursor="hand2",
                activebackground=BG, activeforeground=ACCENT,
                anchor="w", command=command,
            ).pack(fill="x", anchor="w", pady=(0, 4))

        excel_link(excel_footer, "📊  Ouvrir Excel", self._open_excel)
        excel_link(excel_footer, "📁  Changer l'emplacement…", self._open_excel_settings)

    def _current_sort_mode(self) -> SortMode:
        label = self.sort_mode.get()
        for option_label, mode in SORT_OPTIONS:
            if option_label == label:
                return mode
        return DEFAULT_SORT

    def _refresh_project_list(self, *, persist_sort: bool = False) -> None:
        current = self.selected_project.get()
        ordered = sort_projects(self.projects, self._current_sort_mode())
        names = project_names(ordered)
        self.project_picker.set_options(names)
        if current in names:
            self.selected_project.set(current)
        elif names:
            self.selected_project.set(names[0])
        else:
            self.selected_project.set("")
        if persist_sort:
            self._save_config()

    def _save_config(self) -> None:
        save_config({
            "projects": self.projects,
            "sort_mode": self._current_sort_mode(),
            "excel_file": self.excel_file,
        })

    def _on_sort_changed(self, _event=None) -> None:
        self._refresh_project_list(persist_sort=True)

    def _on_root_configure(self, event: tk.Event) -> None:
        if event.widget is self.root:
            self.project_picker.close_if_open()

    # ── Drag ──
    def _start_drag(self, e):
        self._drag_x = e.x
        self._drag_y = e.y

    def _drag(self, e):
        x = self.root.winfo_x() + e.x - self._drag_x
        y = self.root.winfo_y() + e.y - self._drag_y
        self.root.geometry(f"+{x}+{y}")

    # ── Timer ──
    def _start(self):
        project = self.selected_project.get().strip()
        if project not in project_names(self.projects):
            messagebox.showwarning("Projet manquant", "Sélectionnez un projet valide dans la liste.")
            return
        self.tracking   = True
        self.start_time = datetime.now()
        self.elapsed    = 0
        self._running   = True
        self.start_btn.config(state="disabled", bg=self._SUBTLE, fg=self._MUTED)
        self.stop_btn.config(state="normal",   bg=self._RED,    fg="white")
        self.status_label.config(
            text=f"● En cours — {self.selected_project.get()}", fg=self._GREEN)
        threading.Thread(target=self._tick, daemon=True).start()

    def _tick(self):
        while self._running:
            self.elapsed = (datetime.now() - self.start_time).total_seconds()
            h = int(self.elapsed // 3600)
            m = int((self.elapsed % 3600) // 60)
            s = int(self.elapsed % 60)
            self.root.after(0, self.timer_label.config, {"text": f"{h:02}:{m:02}:{s:02}"})
            time.sleep(1)

    def _stop(self):
        if not self.tracking:
            return
        self._running = False
        self.tracking = False
        end_time = datetime.now()
        project  = self.selected_project.get()

        self.start_btn.config(state="normal",   bg=self._GREEN, fg="white")
        self.stop_btn.config(state="disabled",  bg=self._SUBTLE, fg=self._MUTED)
        self.status_label.config(text="✔ Session enregistrée", fg=self._ACCENT)

        notes = simpledialog.askstring(
            "Notes (optionnel)", "Ajouter une note pour cette session :",
            parent=self.root) or ""

        log_session(project, self.start_time, end_time, notes, self.excel_file)

        dur = int(self.elapsed)
        h, m = dur // 3600, (dur % 3600) // 60
        messagebox.showinfo("Session enregistrée",
                            f"✔ {project}\nDurée : {h}h {m:02}min\n\nEnregistré dans :\n{self.excel_file}")
        self.timer_label.config(text="00:00:00")
        self.status_label.config(text="● En attente", fg=self._MUTED)

    # ── Manage projects ──
    def _open_manage(self):
        ManageProjectsDialog(
            self.root,
            self.projects,
            self._current_sort_mode(),
            self.excel_file,
            self._on_projects_updated,
        )

    def _on_projects_updated(self, new_projects: list[ProjectEntry], sort_mode: SortMode):
        self.projects = new_projects
        self.sort_mode.set(SORT_LABELS[sort_mode])
        self._refresh_project_list(persist_sort=True)

    # ── Excel ──
    def _open_excel_settings(self) -> None:
        ExcelFileDialog(self.root, self.excel_file, self._on_excel_file_changed)

    def _on_excel_file_changed(self, excel_path: str) -> None:
        self.excel_file = excel_path
        self._save_config()
        ensure_excel(self.excel_file)
        messagebox.showinfo(
            "Fichier Excel",
            f"Les sessions seront enregistrées dans :\n{self.excel_file}",
            parent=self.root,
        )

    def _open_excel(self) -> None:
        open_excel_file(self.excel_file, parent=self.root)

# ─── ENTRY ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ensure_installed()
    startup_config = load_config()
    ensure_excel(startup_config["excel_file"])
    save_config(startup_config)
    root = tk.Tk()
    app  = TimeTrackerApp(root)
    root.mainloop()