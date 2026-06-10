import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import time
import threading
import hashlib
from datetime import datetime, timedelta
from typing import Literal, TypedDict
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

DEFAULT_SORT: SortMode = "alpha_asc"
SORT_OPTIONS: list[tuple[str, SortMode]] = [
    ("A → Z", "alpha_asc"),
    ("Date ↑", "date_asc"),
    ("Date ↓", "date_desc"),
]
SORT_LABELS: dict[SortMode, str] = {mode: label for label, mode in SORT_OPTIONS}
APP_VERSION = "1.1.0"

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
EXCEL_FILE  = os.path.join(APP_DIR, "heures_projet.xlsx")
CONFIG_FILE = os.path.join(APP_DIR, "projects.json")

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
            return {"projects": _migrate_legacy_projects(raw), "sort_mode": DEFAULT_SORT}
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
        return {
            "projects": normalized,
            "sort_mode": _normalize_sort_mode(str(raw.get("sort_mode", DEFAULT_SORT))),
        }
    return {"projects": _default_projects(), "sort_mode": DEFAULT_SORT}

def save_config(projects: list[ProjectEntry], sort_mode: SortMode) -> None:
    payload: ProjectsConfig = {"projects": projects, "sort_mode": sort_mode}
    with open(CONFIG_FILE, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)

def sort_projects(projects: list[ProjectEntry], mode: SortMode) -> list[ProjectEntry]:
    if mode == "alpha_asc":
        return sorted(projects, key=lambda entry: entry["name"].casefold())
    if mode == "date_asc":
        return sorted(projects, key=lambda entry: entry["created"])
    return sorted(projects, key=lambda entry: entry["created"], reverse=True)

def filter_projects(projects: list[ProjectEntry], query: str) -> list[ProjectEntry]:
    needle = query.strip().casefold()
    if not needle:
        return projects
    return [entry for entry in projects if needle in entry["name"].casefold()]

def project_names(projects: list[ProjectEntry]) -> list[str]:
    return [entry["name"] for entry in projects]

class FilterBar(tk.Frame):
    """Barre recherche (placeholder + effacer) et tri dans un seul bloc."""

    PLACEHOLDER = "Filtrer les projets…"

    def __init__(
        self,
        parent: tk.Misc,
        search_var: tk.StringVar,
        sort_var: tk.StringVar,
        on_filter,
        on_sort,
        *,
        bg: str,
        field_bg: str,
        border: str,
        muted: str,
        accent: str,
        text: str,
        sort_width: int = 9,
        combobox_style: str | None = None,
    ) -> None:
        super().__init__(parent, bg=bg)
        self.search_var = search_var
        self._on_filter = on_filter
        self._muted = muted
        self._accent = accent
        self._field_bg = field_bg
        self._text_fg = text
        self._placeholder_active = False
        self._updating = False

        shell = tk.Frame(self, bg=border, padx=1, pady=1)
        shell.pack(fill="x")
        inner = tk.Frame(shell, bg=field_bg)
        inner.pack(fill="x")

        tk.Label(inner, text="⌕", bg=field_bg, fg=muted,
                 font=("Segoe UI", 11)).pack(side="left", padx=(10, 4))

        self._entry_var = tk.StringVar()
        self.entry = tk.Entry(
            inner, textvariable=self._entry_var,
            font=("Segoe UI", 10), bg=field_bg, fg=muted,
            relief="flat", bd=0, highlightthickness=0,
        )
        self.entry.pack(side="left", fill="x", expand=True, ipady=6)

        self.clear_btn = tk.Button(
            inner, text="×", bg=field_bg, fg=muted, bd=0,
            font=("Segoe UI", 12), cursor="hand2", padx=4,
            activebackground=field_bg, activeforeground=accent,
            command=self.clear,
        )
        self.clear_btn.pack(side="right", padx=(0, 2))

        self._separator = tk.Frame(inner, bg=border, width=1)
        self._separator.pack(side="right", fill="y", padx=(4, 6), pady=6)

        sort_kwargs: dict[str, object] = {
            "textvariable": sort_var,
            "state": "readonly",
            "values": [label for label, _mode in SORT_OPTIONS],
            "font": ("Segoe UI", 9),
            "width": sort_width,
        }
        if combobox_style:
            sort_kwargs["style"] = combobox_style
        sort_combo = ttk.Combobox(inner, **sort_kwargs)
        sort_combo.pack(side="right", padx=(0, 8), ipady=2)
        sort_combo.bind("<<ComboboxSelected>>", on_sort)

        self.entry.bind("<FocusIn>", self._on_focus_in)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.entry.bind("<KeyPress>", self._on_key_press)
        self._entry_var.trace_add("write", self._on_text_changed)
        self._show_placeholder()

    def _set_entry_text(self, value: str) -> None:
        self._updating = True
        self._entry_var.set(value)
        self._updating = False

    def _show_placeholder(self) -> None:
        self._placeholder_active = True
        self._set_entry_text(self.PLACEHOLDER)
        self.entry.config(fg=self._muted)

    def _leave_placeholder(self) -> None:
        if not self._placeholder_active:
            return
        self._placeholder_active = False
        self._set_entry_text("")
        self.entry.config(fg=self._text_fg)

    def _on_focus_in(self, _event: tk.Event) -> None:
        if self._placeholder_active:
            self._leave_placeholder()

    def _on_focus_out(self, _event: tk.Event) -> None:
        if self._placeholder_active:
            return
        if not self._entry_var.get().strip():
            self.search_var.set("")
            self._show_placeholder()
            self._on_filter()

    def _on_key_press(self, event: tk.Event) -> None:
        if self._placeholder_active and event.char:
            self._leave_placeholder()

    def _on_text_changed(self, *_args: str) -> None:
        if self._updating or self._placeholder_active:
            return
        self.search_var.set(self._entry_var.get().strip())
        self._on_filter()

    def clear(self) -> None:
        self._placeholder_active = False
        self._set_entry_text("")
        self.search_var.set("")
        self.entry.config(fg=self._text_fg)
        self.entry.focus_set()
        self._on_filter()

    def reset(self) -> None:
        self.clear()

# ─── EXCEL ─────────────────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", fgColor="2D6A4F")
SUBHEAD_FILL = PatternFill("solid", fgColor="40916C")
ALT_FILL     = PatternFill("solid", fgColor="F0FFF4")
WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")

def thin_border():
    s = Side(style="thin", color="B7E4C7")
    return Border(left=s, right=s, top=s, bottom=s)

def ensure_excel():
    if os.path.exists(EXCEL_FILE):
        return
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
    wb.save(EXCEL_FILE)

def log_session(project, start_dt, end_dt, notes=""):
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Sessions"]
    duration_h   = (end_dt - start_dt).total_seconds() / 3600
    duration_min = (end_dt - start_dt).total_seconds() / 60
    nr   = ws.max_row + 1
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
    wb.save(EXCEL_FILE)

def rename_project_in_excel(old_name, new_name):
    if not os.path.exists(EXCEL_FILE):
        return
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Sessions"]
    for row in ws.iter_rows(min_row=2):
        if row[1].value == old_name:
            row[1].value = new_name
    _refresh_summary(wb)
    wb.save(EXCEL_FILE)

def _refresh_summary(wb):
    ws  = wb["Sessions"]
    ws2 = wb["Résumé par projet"]
    for row in ws2.iter_rows(min_row=3, max_row=ws2.max_row):
        for cell in row:
            cell.value = None
            cell.fill  = WHITE_FILL
    totals = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1]:
            p = row[1]
            h = row[4] or 0
            totals[p] = totals.get(p, {"h": 0, "n": 0})
            totals[p]["h"] += h
            totals[p]["n"] += 1
    for i, (p, data) in enumerate(sorted(totals.items()), start=3):
        fill2 = ALT_FILL if i % 2 == 0 else WHITE_FILL
        for col, val in enumerate([p, round(data["h"], 2), data["n"]], 1):
            cell = ws2.cell(row=i, column=col, value=val)
            cell.font      = Font(name="Calibri", size=10)
            cell.fill      = fill2
            cell.border    = thin_border()
            cell.alignment = Alignment(horizontal="center" if col != 1 else "left")
        ws2.cell(row=i, column=2).number_format = "0.00"

# ─── MANAGE PROJECTS DIALOG ────────────────────────────────────────────────────

class ManageProjectsDialog(tk.Toplevel):
    """Fenêtre modale pour gérer (modifier / supprimer) les projets."""

    def __init__(
        self,
        parent: tk.Misc,
        projects: list[ProjectEntry],
        sort_mode: SortMode,
        on_close,
    ):
        super().__init__(parent)
        self.projects: list[ProjectEntry] = list(projects)
        self.sort_mode = tk.StringVar(value=SORT_LABELS[sort_mode])
        self.search_var = tk.StringVar()
        self.on_close = on_close
        self.title("Gérer les projets")
        self.resizable(False, False)
        self.grab_set()
        self.attributes("-topmost", True)
        self.configure(bg="#FFFFFF")
        self._build()
        self._refresh_list()
        self.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        self.geometry(f"+{px - 260}+{py}")

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

        tk.Label(body, text="Vos projets", bg=BG, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

        self.filter_bar = FilterBar(
            body, self.search_var, self.sort_mode,
            on_filter=lambda: self._refresh_list(),
            on_sort=self._on_sort_selected,
            bg=BG, field_bg="#F5F5F5", border=BORDER,
            muted=MUTED, accent=GREEN, text="#222222",
        )
        self.filter_bar.pack(fill="x", pady=(0, 4))

        self.count_label = tk.Label(body, text="", bg=BG, fg=MUTED,
                                    font=("Segoe UI", 8))
        self.count_label.pack(anchor="w", pady=(0, 6))

        list_frame = tk.Frame(body, bg=BORDER, bd=1, relief="flat")
        list_frame.pack(fill="both", expand=True)
        sb = tk.Scrollbar(list_frame, orient="vertical")
        self.lb = tk.Listbox(
            list_frame, yscrollcommand=sb.set, selectmode="single",
            font=("Segoe UI", 10), bg=BG, fg="#222222",
            selectbackground=GREEN, selectforeground="white",
            activestyle="none", relief="flat", bd=0,
            highlightthickness=0, width=28, height=10,
        )
        sb.config(command=self.lb.yview)
        sb.pack(side="right", fill="y")
        self.lb.pack(fill="both", expand=True, padx=2, pady=2)

        btn_frame = tk.Frame(body, bg=BG)
        btn_frame.pack(fill="x", pady=(10, 0))

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

    def _on_sort_selected(self, _event=None) -> None:
        self._refresh_list()

    def _visible_projects(self) -> list[ProjectEntry]:
        ordered = sort_projects(self.projects, self._current_sort_mode())
        return filter_projects(ordered, self.search_var.get())

    def _refresh_list(self) -> None:
        selected = self._selected_name()
        self.lb.delete(0, "end")
        visible = self._visible_projects()
        for entry in visible:
            self.lb.insert("end", entry["name"])

        total = len(self.projects)
        shown = len(visible)
        query = self.search_var.get().strip()
        if query:
            suffix = "trouvé" if shown == 1 else "trouvés"
            self.count_label.config(text=f"{shown} sur {total} {suffix}")
        else:
            label = "projet" if total == 1 else "projets"
            self.count_label.config(text=f"{total} {label}")

        if selected:
            for index, entry in enumerate(visible):
                if entry["name"] == selected:
                    self.lb.selection_set(index)
                    self.lb.see(index)
                    break
        elif visible and query:
            self.lb.selection_set(0)
            self.lb.see(0)

    def _selected_name(self) -> str | None:
        sel = self.lb.curselection()
        if not sel:
            return None
        return self.lb.get(sel[0])

    def _require_selection(self) -> str | None:
        name = self._selected_name()
        if not name:
            messagebox.showinfo("Sélection", "Sélectionnez d'abord un projet dans la liste.", parent=self)
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
        self._refresh_list()
        rename_project_in_excel(old_name, new_name)

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
        self._refresh_list()

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
        self.filter_bar.reset()
        self._refresh_list()
        visible = self._visible_projects()
        for index, entry in enumerate(visible):
            if entry["name"] == name:
                self.lb.selection_set(index)
                self.lb.see(index)
                break

    def _close(self):
        self.destroy()

    def destroy(self):
        save_config(self.projects, self._current_sort_mode())
        self.on_close(self.projects, self._current_sort_mode())
        self.grab_release()
        super().destroy()

# ─── MAIN APP ──────────────────────────────────────────────────────────────────

class TimeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("⏱ Time Tracker")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.98)

        config = load_config()
        self.projects: list[ProjectEntry] = config["projects"]
        self.sort_mode = tk.StringVar(value=SORT_LABELS[config["sort_mode"]])
        self.search_var = tk.StringVar()
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
        tk.Button(top_row, text="Gérer ›", bg=BG, fg=GREEN, bd=0,
                  font=("Segoe UI", 7, "underline"), cursor="hand2",
                  activebackground=BG, activeforeground=ACCENT,
                  command=self._open_manage).pack(side="right")

        style = ttk.Style()
        style.theme_use("default")
        style.configure("W.TCombobox",
                        fieldbackground=SUBTLE, background=SUBTLE,
                        foreground=TEXT, selectbackground=GREEN,
                        selectforeground="white", arrowcolor=ACCENT,
                        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
        style.map("W.TCombobox", fieldbackground=[("readonly", SUBTLE)])

        self.filter_bar = FilterBar(
            body, self.search_var, self.sort_mode,
            on_filter=lambda: self._refresh_project_list(),
            on_sort=self._on_sort_changed,
            bg=BG, field_bg=SUBTLE, border=BORDER,
            muted=MUTED, accent=GREEN, text=TEXT,
            combobox_style="W.TCombobox", sort_width=8,
        )
        self.filter_bar.pack(fill="x", pady=(0, 6))

        combo_shell = tk.Frame(body, bg=BORDER, padx=1, pady=1)
        combo_shell.pack(fill="x", pady=(0, 12))
        combo_inner = tk.Frame(combo_shell, bg=SUBTLE)
        combo_inner.pack(fill="x")

        self.combo = ttk.Combobox(
            combo_inner, textvariable=self.selected_project,
            values=[], state="readonly",
            font=("Segoe UI", 10), style="W.TCombobox",
        )
        self.combo.pack(fill="x", padx=8, ipady=4)

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

        tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=(12, 8))

        tk.Button(body, text="📊  Ouvrir le fichier Excel", bg=BG, fg=MUTED,
                  font=("Segoe UI", 8), bd=0, cursor="hand2",
                  activebackground=BG, activeforeground=ACCENT,
                  command=self._open_excel).pack()

    def _current_sort_mode(self) -> SortMode:
        label = self.sort_mode.get()
        for option_label, mode in SORT_OPTIONS:
            if option_label == label:
                return mode
        return DEFAULT_SORT

    def _refresh_project_list(self, *, persist_sort: bool = False) -> None:
        current = self.selected_project.get()
        ordered = sort_projects(self.projects, self._current_sort_mode())
        visible = filter_projects(ordered, self.search_var.get())
        names = project_names(visible)
        self.combo["values"] = names
        if current in names:
            self.selected_project.set(current)
        elif names:
            self.selected_project.set(names[0])
        else:
            self.selected_project.set("")
        if persist_sort:
            save_config(self.projects, self._current_sort_mode())

    def _on_sort_changed(self, _event=None) -> None:
        self._refresh_project_list(persist_sort=True)

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

        log_session(project, self.start_time, end_time, notes)

        dur = int(self.elapsed)
        h, m = dur // 3600, (dur % 3600) // 60
        messagebox.showinfo("Session enregistrée",
                            f"✔ {project}\nDurée : {h}h {m:02}min\n\nEnregistré dans :\n{EXCEL_FILE}")
        self.timer_label.config(text="00:00:00")
        self.status_label.config(text="● En attente", fg=self._MUTED)

    # ── Manage projects ──
    def _open_manage(self):
        ManageProjectsDialog(
            self.root,
            self.projects,
            self._current_sort_mode(),
            self._on_projects_updated,
        )

    def _on_projects_updated(self, new_projects: list[ProjectEntry], sort_mode: SortMode):
        self.projects = new_projects
        self.sort_mode.set(SORT_LABELS[sort_mode])
        self._refresh_project_list(persist_sort=True)

    # ── Excel ──
    def _open_excel(self):
        ensure_excel()
        if os.name == "nt":
            os.startfile(EXCEL_FILE)
        else:
            try:
                sysname = os.uname().sysname
            except AttributeError:
                sysname = ""
            if sysname == "Darwin":
                os.system(f'open "{EXCEL_FILE}"')
            else:
                os.system(f'xdg-open "{EXCEL_FILE}"')

# ─── ENTRY ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ensure_installed()
    ensure_excel()
    startup_config = load_config()
    save_config(startup_config["projects"], startup_config["sort_mode"])
    root = tk.Tk()
    app  = TimeTrackerApp(root)
    root.mainloop()