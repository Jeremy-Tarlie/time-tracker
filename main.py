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
APP_VERSION = "1.1.3"

INSTALL_DIR   = os.path.join(os.environ.get("LOCALAPPDATA", ""), "TimeTracker")
INSTALL_EXE   = os.path.join(INSTALL_DIR, "TimeTracker.exe")
REGISTRY_KEY  = r"Software\Microsoft\Windows\CurrentVersion\Run"
REGISTRY_NAME = "TimeTracker"
APP_NAME      = "Time Tracker"

class ThemeColors(TypedDict):
    bg: str
    surface: str
    border: str
    green: str
    green_dark: str
    green_soft: str
    timer: str
    text: str
    muted: str
    danger: str
    danger_soft: str
    danger_hover: str
    white: str

THEME: ThemeColors = {
    "bg": "#FFFFFF",
    "surface": "#F7F9F8",
    "border": "#DCE5E1",
    "green": "#40916C",
    "green_dark": "#2D6A4F",
    "green_soft": "#D8F3DC",
    "timer": "#1B4332",
    "text": "#1A1A1A",
    "muted": "#5C6B66",
    "danger": "#E63946",
    "danger_soft": "#FDECEE",
    "danger_hover": "#C1121F",
    "white": "#FFFFFF",
}

THEME_REQUIRED_KEYS: tuple[str, ...] = (
    "bg", "surface", "border", "green", "green_dark", "green_soft",
    "timer", "text", "muted", "danger", "danger_soft", "danger_hover", "white",
)

_BUTTON_HOVER_NORMAL: dict[int, str] = {}
_BUTTON_HOVER_ACTIVE: dict[int, str] = {}

def _bind_button_hover(button: tk.Button, *, normal_bg: str, hover_bg: str) -> None:
    key = id(button)
    _BUTTON_HOVER_NORMAL[key] = normal_bg
    _BUTTON_HOVER_ACTIVE[key] = hover_bg

    def on_enter(_event: tk.Event) -> None:
        if str(button.cget("state")) == "disabled":
            return
        button.configure(bg=_BUTTON_HOVER_ACTIVE[key])

    def on_leave(_event: tk.Event) -> None:
        # Toujours restaurer la couleur « normale » (y compris si disabled entre Enter et Leave).
        button.configure(bg=_BUTTON_HOVER_NORMAL[key])

    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)

def set_button_palette(
    button: tk.Button,
    *,
    bg: str,
    fg: str,
    hover_bg: str,
    state: str | None = None,
) -> None:
    kwargs: dict[str, str] = {
        "bg": bg,
        "fg": fg,
        "activebackground": hover_bg,
        "activeforeground": fg,
    }
    if state is not None:
        kwargs["state"] = state
    button.configure(**kwargs)
    _BUTTON_HOVER_NORMAL[id(button)] = bg
    _BUTTON_HOVER_ACTIVE[id(button)] = hover_bg

def make_header(parent: tk.Misc, title: str) -> tk.Frame:
    hdr = tk.Frame(parent, bg=THEME["green_dark"], height=38)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    tk.Label(
        hdr,
        text=title,
        bg=THEME["green_dark"],
        fg=THEME["white"],
        font=("Segoe UI", 11, "bold"),
    ).pack(side="left", padx=14, pady=8)
    return hdr

def primary_button(
    parent: tk.Misc,
    text: str,
    command: Callable[[], None],
    *,
    padx: int = 12,
    pady: int = 7,
    font: tuple[str, int, str] = ("Segoe UI", 9, "bold"),
) -> tk.Button:
    btn = tk.Button(
        parent,
        text=text,
        bg=THEME["green"],
        fg=THEME["white"],
        font=font,
        bd=0,
        cursor="hand2",
        padx=padx,
        pady=pady,
        activebackground=THEME["green_dark"],
        activeforeground=THEME["white"],
        disabledforeground=THEME["muted"],
        relief="flat",
        highlightthickness=0,
        command=command,
    )
    _bind_button_hover(btn, normal_bg=THEME["green"], hover_bg=THEME["green_dark"])
    return btn

def secondary_button(
    parent: tk.Misc,
    text: str,
    command: Callable[[], None],
    *,
    padx: int = 10,
    pady: int = 6,
) -> tk.Button:
    btn = tk.Button(
        parent,
        text=text,
        bg=THEME["surface"],
        fg=THEME["text"],
        font=("Segoe UI", 9),
        bd=0,
        cursor="hand2",
        padx=padx,
        pady=pady,
        activebackground=THEME["green_soft"],
        activeforeground=THEME["green_dark"],
        relief="flat",
        highlightthickness=0,
        command=command,
    )
    _bind_button_hover(btn, normal_bg=THEME["surface"], hover_bg=THEME["green_soft"])
    return btn

def ghost_button(
    parent: tk.Misc,
    text: str,
    command: Callable[[], None],
    *,
    fg: str | None = None,
    font: tuple[str, int] | tuple[str, int, str] = ("Segoe UI", 8),
    underline: bool = False,
) -> tk.Button:
    color = fg if fg is not None else THEME["muted"]
    resolved_font: tuple[str, int] | tuple[str, int, str]
    if underline and len(font) == 2:
        resolved_font = (font[0], font[1], "underline")
    else:
        resolved_font = font
    btn = tk.Button(
        parent,
        text=text,
        bg=THEME["bg"],
        fg=color,
        font=resolved_font,
        bd=0,
        cursor="hand2",
        activebackground=THEME["green_soft"],
        activeforeground=THEME["green_dark"],
        relief="flat",
        highlightthickness=0,
        anchor="w",
        command=command,
    )
    _bind_button_hover(btn, normal_bg=THEME["bg"], hover_bg=THEME["green_soft"])
    return btn

def danger_button(
    parent: tk.Misc,
    text: str,
    command: Callable[[], None],
    *,
    padx: int = 10,
    pady: int = 6,
) -> tk.Button:
    btn = tk.Button(
        parent,
        text=text,
        bg=THEME["danger"],
        fg=THEME["white"],
        font=("Segoe UI", 9, "bold"),
        bd=0,
        cursor="hand2",
        padx=padx,
        pady=pady,
        activebackground=THEME["danger_hover"],
        activeforeground=THEME["white"],
        relief="flat",
        highlightthickness=0,
        command=command,
    )
    _bind_button_hover(btn, normal_bg=THEME["danger"], hover_bg=THEME["danger_hover"])
    return btn

def themed_entry(parent: tk.Misc, textvariable: tk.StringVar, *, font: tuple[str, int] = ("Segoe UI", 9)) -> tk.Entry:
    return tk.Entry(
        parent,
        textvariable=textvariable,
        font=font,
        bg=THEME["surface"],
        fg=THEME["text"],
        relief="flat",
        bd=0,
        highlightthickness=1,
        highlightbackground=THEME["border"],
        highlightcolor=THEME["green"],
        insertbackground=THEME["text"],
    )

_SORT_COMBO_STYLE: str | None = None

def configure_sort_combobox_style() -> str:
    global _SORT_COMBO_STYLE
    if _SORT_COMBO_STYLE is not None:
        return _SORT_COMBO_STYLE
    style = ttk.Style()
    style.theme_use("default")
    style_name = "Theme.TCombobox"
    style.configure(
        style_name,
        fieldbackground=THEME["surface"],
        background=THEME["surface"],
        foreground=THEME["text"],
        selectbackground=THEME["green"],
        selectforeground=THEME["white"],
        arrowcolor=THEME["green_dark"],
        bordercolor=THEME["border"],
        lightcolor=THEME["border"],
        darkcolor=THEME["border"],
    )
    style.map(style_name, fieldbackground=[("readonly", THEME["surface"])])
    _SORT_COMBO_STYLE = style_name
    return style_name

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
        hover_bg: str,
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
        self._hover_bg = hover_bg
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
            widget.bind("<Enter>", self._on_trigger_enter)
            widget.bind("<Leave>", self._on_trigger_leave)

        self.variable.trace_add("write", lambda *_: self._update_display())
        self._update_display()

    def _on_trigger_enter(self, _event: tk.Event) -> None:
        if not self._open:
            self._set_trigger_bg(self._hover_bg)

    def _on_trigger_leave(self, _event: tk.Event) -> None:
        if not self._open:
            self._set_trigger_bg(self._field_bg)

    def _set_trigger_bg(self, color: str) -> None:
        self.trigger.configure(bg=color)
        for child in self.trigger.winfo_children():
            if isinstance(child, tk.Label):
                child.configure(bg=color)

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
        self._set_trigger_bg(self._field_bg)

# ─── EXCEL ─────────────────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", fgColor="2D6A4F")
SUBHEAD_FILL = PatternFill("solid", fgColor="40916C")
ALT_FILL     = PatternFill("solid", fgColor="F0FFF4")
WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")

def thin_border():
    s = Side(style="thin", color="B7E4C7")
    return Border(left=s, right=s, top=s, bottom=s)

class ExcelWriteError(Exception):
    """Fichier Excel inaccessible en écriture (ouvert ailleurs, verrouillé, etc.)."""

EXCEL_WRITE_HINT = (
    "Fermez le fichier Excel/Calc, enregistrez vos modifications, puis réessayez."
)

OOXML_ZIP_MAGIC = b"PK"
OLE_CFBF_MAGIC = b"\xd0\xcf\x11\xe0"

def _row_has_data(values: tuple[object, ...]) -> bool:
    return any(value is not None and str(value).strip() != "" for value in values)

def _session_row_values(ws, row_idx: int) -> tuple[object, ...]:
    return tuple(ws.cell(row=row_idx, column=col).value for col in range(1, 8))

def _last_row_with_data(ws, *, min_row: int = 2) -> int:
    """Dernière ligne contenant des données (openpyxl.max_row peut être incorrect après édition manuelle)."""
    last = min_row - 1
    for row_idx, row in enumerate(ws.iter_rows(min_row=min_row, values_only=True), start=min_row):
        if _row_has_data(row):
            last = row_idx
    return last

def _next_session_row(ws) -> int:
    """Prochaine ligne vide après la dernière session (ne jamais écraser une ligne remplie)."""
    last = _last_row_with_data(ws, min_row=2)
    nr = max(last + 1, 2)
    while _row_has_data(_session_row_values(ws, nr)):
        nr += 1
    return nr

def _excel_lock_paths(excel_path: str) -> list[str]:
    directory, filename = os.path.split(excel_path)
    return [
        os.path.join(directory, f"~${filename}"),           # Microsoft Excel
        os.path.join(directory, f".~lock.{filename}#"),     # LibreOffice / OpenOffice
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

def _save_workbook_atomic(wb, excel_path: str) -> None:
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
    _save_workbook_atomic(wb, excel_path)

def log_session(project: str, start_dt: datetime, end_dt: datetime, notes: str, excel_path: str) -> int:
    """Ajoute une session en append. Retourne le numéro de ligne écrite. Lève ExcelWriteError si lock."""
    ensure_excel(excel_path)
    _assert_excel_writable(excel_path)
    try:
        wb = load_workbook(excel_path)
    except ExcelWriteError:
        raise
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
    duration_h   = (end_dt - start_dt).total_seconds() / 3600
    duration_min = (end_dt - start_dt).total_seconds() / 60
    nr   = _next_session_row(ws)
    if _row_has_data(_session_row_values(ws, nr)):
        raise ExcelWriteError(
            f"La ligne {nr} de la feuille Sessions n'est pas vide.\n"
            "Aucune session n'a été écrasée."
        )
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
    _save_workbook_atomic(wb, excel_path)
    return nr

def rename_project_in_excel(old_name: str, new_name: str, excel_path: str) -> None:
    if not os.path.exists(excel_path):
        return
    _assert_excel_writable(excel_path)
    wb = load_workbook(excel_path)
    ws = wb["Sessions"]
    for row in ws.iter_rows(min_row=2):
        if row[1].value == old_name:
            row[1].value = new_name
    _refresh_summary(wb)
    _save_workbook_atomic(wb, excel_path)

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
        self.configure(bg=THEME["bg"])
        self._build()
        self.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        self.geometry(f"500x220+{px - 250}+{py}")

    def _build(self) -> None:
        make_header(self, "Fichier Excel")

        body = tk.Frame(self, bg=THEME["bg"], padx=16, pady=14)
        body.pack(fill="both", expand=True)

        tk.Label(
            body,
            text="EMPLACEMENT DU FICHIER D'EXPORT",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=("Segoe UI", 7, "bold"),
        ).pack(anchor="w", pady=(0, 6))

        path_row = tk.Frame(body, bg=THEME["bg"])
        path_row.pack(fill="x", pady=(0, 8))
        self._path_entry = themed_entry(path_row, self._path_var)
        self._path_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 6))
        secondary_button(path_row, "Ouvrir", self._open).pack(side="left", padx=(0, 6))
        primary_button(path_row, "Parcourir…", self._browse, padx=10, pady=6).pack(side="left")

        tk.Label(
            body,
            text="Les nouvelles sessions seront enregistrées à cet emplacement.\n"
                 "L'ancien fichier n'est pas déplacé automatiquement.",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=("Segoe UI", 8),
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        btn_row = tk.Frame(body, bg=THEME["bg"])
        btn_row.pack(fill="x")

        secondary_button(btn_row, "Par défaut", self._use_default).pack(side="left")
        secondary_button(btn_row, "Annuler", self.destroy).pack(side="right", padx=(6, 0))
        primary_button(btn_row, "Enregistrer", self._save).pack(side="right")

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
        self.configure(bg=THEME["bg"])
        self._build()
        self._refresh_picker()
        self.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        self.geometry(f"320x280+{px - 260}+{py}")

    def _build(self) -> None:
        make_header(self, "Gérer les projets")

        body = tk.Frame(self, bg=THEME["bg"], padx=16, pady=14)
        body.pack(fill="both", expand=True)

        top_row = tk.Frame(body, bg=THEME["bg"])
        top_row.pack(fill="x", pady=(0, 6))
        tk.Label(
            top_row,
            text="PROJET",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=("Segoe UI", 7, "bold"),
        ).pack(side="left")

        sort_combo = ttk.Combobox(
            top_row,
            textvariable=self.sort_mode,
            state="readonly",
            values=[label for label, _mode in SORT_OPTIONS],
            font=("Segoe UI", 8),
            width=8,
            style=configure_sort_combobox_style(),
        )
        sort_combo.pack(side="right", padx=(8, 0))
        sort_combo.bind("<<ComboboxSelected>>", self._on_sort_selected)

        self.project_picker = SearchableDropdown(
            body,
            self.selected_name,
            bg=THEME["bg"],
            field_bg=THEME["surface"],
            border=THEME["border"],
            muted=THEME["muted"],
            accent=THEME["green_dark"],
            highlight=THEME["green"],
            text=THEME["text"],
            hover_bg=THEME["green_soft"],
            placeholder="Choisir un projet…",
        )
        self.project_picker.pack(fill="x", pady=(0, 12))
        self.bind("<Configure>", self._on_dialog_configure)

        btn_frame = tk.Frame(body, bg=THEME["bg"])
        btn_frame.pack(fill="x")

        primary_button(btn_frame, "Renommer", self._rename, padx=10, pady=6).pack(
            side="left", padx=(0, 6),
        )
        danger_button(btn_frame, "Supprimer", self._delete).pack(side="left")

        tk.Frame(body, bg=THEME["border"], height=1).pack(fill="x", pady=(14, 10))

        tk.Label(
            body,
            text="AJOUTER UN PROJET",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=("Segoe UI", 7, "bold"),
        ).pack(anchor="w", pady=(0, 4))
        add_row = tk.Frame(body, bg=THEME["bg"])
        add_row.pack(fill="x")
        self.new_var = tk.StringVar()
        entry = themed_entry(add_row, self.new_var, font=("Segoe UI", 10))
        entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 6))
        entry.bind("<Return>", lambda _event: self._add())
        primary_button(add_row, "Ajouter", self._add, padx=10, pady=6).pack(side="left")

        secondary_button(body, "Fermer", self._close).pack(anchor="e", pady=(14, 0))

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
        self._timer_gen       = 0

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

    def _build_ui(self) -> None:
        self.root.configure(bg=THEME["bg"])

        title_bar = make_header(self.root, "Time Tracker")
        for widget in [title_bar] + title_bar.winfo_children():
            if not isinstance(widget, tk.Button):
                widget.bind("<ButtonPress-1>", self._start_drag)
                widget.bind("<B1-Motion>", self._drag)

        body = tk.Frame(self.root, bg=THEME["bg"], padx=16, pady=16)
        body.pack(fill="both", expand=True)

        top_row = tk.Frame(body, bg=THEME["bg"])
        top_row.pack(fill="x", pady=(0, 6))
        tk.Label(
            top_row,
            text="PROJET",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=("Segoe UI", 7, "bold"),
        ).pack(side="left")

        sort_combo = ttk.Combobox(
            top_row,
            textvariable=self.sort_mode,
            state="readonly",
            values=[label for label, _mode in SORT_OPTIONS],
            font=("Segoe UI", 8),
            width=8,
            style=configure_sort_combobox_style(),
        )
        sort_combo.pack(side="right", padx=(8, 0))
        sort_combo.bind("<<ComboboxSelected>>", self._on_sort_changed)

        ghost_button(
            top_row,
            "Gérer ›",
            self._open_manage,
            fg=THEME["green"],
            font=("Segoe UI", 7),
            underline=True,
        ).pack(side="right")

        self.project_picker = SearchableDropdown(
            body,
            self.selected_project,
            bg=THEME["bg"],
            field_bg=THEME["surface"],
            border=THEME["border"],
            muted=THEME["muted"],
            accent=THEME["green_dark"],
            highlight=THEME["green"],
            text=THEME["text"],
            hover_bg=THEME["green_soft"],
        )
        self.project_picker.pack(fill="x", pady=(0, 14))
        self.root.bind("<Configure>", self._on_root_configure)

        tk.Frame(body, bg=THEME["border"], height=1).pack(fill="x", pady=(0, 14))

        self.timer_label = tk.Label(
            body,
            text="00:00:00",
            bg=THEME["bg"],
            fg=THEME["timer"],
            font=("Segoe UI Mono", 32, "bold"),
        )
        self.timer_label.pack()

        self.status_label = tk.Label(
            body,
            text="● En attente",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=("Segoe UI", 9),
        )
        self.status_label.pack(pady=(4, 16))

        btn_row = tk.Frame(body, bg=THEME["bg"])
        btn_row.pack(fill="x")

        self.start_btn = primary_button(
            btn_row,
            "▶  Start",
            self._start,
            padx=14,
            pady=10,
            font=("Segoe UI", 10, "bold"),
        )
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.stop_btn = tk.Button(
            btn_row,
            text="■  Stop",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            cursor="hand2",
            padx=14,
            pady=10,
            disabledforeground=THEME["muted"],
            relief="flat",
            highlightthickness=0,
            command=self._stop,
        )
        _bind_button_hover(
            self.stop_btn,
            normal_bg=THEME["surface"],
            hover_bg=THEME["danger_soft"],
        )
        set_button_palette(
            self.stop_btn,
            bg=THEME["surface"],
            fg=THEME["muted"],
            hover_bg=THEME["danger_soft"],
            state="disabled",
        )
        self.stop_btn.pack(side="left", fill="x", expand=True)

        tk.Frame(body, bg=THEME["border"], height=1).pack(fill="x", pady=(14, 12))

        excel_footer = tk.Frame(body, bg=THEME["bg"])
        excel_footer.pack(fill="x", pady=(0, 4))

        ghost_button(excel_footer, "Ouvrir Excel", self._open_excel).pack(
            fill="x", anchor="w", pady=(0, 4),
        )
        ghost_button(
            excel_footer,
            "Changer l'emplacement…",
            self._open_excel_settings,
        ).pack(fill="x", anchor="w")

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
    def _set_idle_controls(self) -> None:
        set_button_palette(
            self.start_btn,
            bg=THEME["green"],
            fg=THEME["white"],
            hover_bg=THEME["green_dark"],
            state="normal",
        )
        set_button_palette(
            self.stop_btn,
            bg=THEME["surface"],
            fg=THEME["muted"],
            hover_bg=THEME["danger_soft"],
            state="disabled",
        )

    def _set_busy_controls(self) -> None:
        set_button_palette(
            self.start_btn,
            bg=THEME["surface"],
            fg=THEME["muted"],
            hover_bg=THEME["surface"],
            state="disabled",
        )
        set_button_palette(
            self.stop_btn,
            bg=THEME["surface"],
            fg=THEME["muted"],
            hover_bg=THEME["surface"],
            state="disabled",
        )

    def _ask_session_notes(self) -> str:
        dialog = tk.Toplevel(self.root)
        dialog.title("Notes (optionnel)")
        dialog.configure(bg=THEME["bg"])
        dialog.attributes("-topmost", True)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        result: dict[str, str] = {"value": ""}

        body = tk.Frame(dialog, bg=THEME["bg"], padx=16, pady=14)
        body.pack(fill="both", expand=True)
        tk.Label(
            body,
            text="Ajouter une note pour cette session :",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(0, 8))
        note_var = tk.StringVar()
        entry = themed_entry(body, note_var, font=("Segoe UI", 10))
        entry.pack(fill="x", ipady=6)
        entry.focus_set()

        def confirm() -> None:
            result["value"] = note_var.get().strip()
            dialog.destroy()

        def skip() -> None:
            result["value"] = ""
            dialog.destroy()

        btn_row = tk.Frame(body, bg=THEME["bg"])
        btn_row.pack(fill="x", pady=(12, 0))
        secondary_button(btn_row, "Sans note", skip).pack(side="left")
        primary_button(btn_row, "OK", confirm).pack(side="right")
        entry.bind("<Return>", lambda _event: confirm())
        dialog.bind("<Escape>", lambda _event: skip())

        dialog.update_idletasks()
        px = self.root.winfo_rootx() + max((self.root.winfo_width() - 320) // 2, 0)
        py = self.root.winfo_rooty() + 40
        dialog.geometry(f"320x150+{px}+{py}")
        dialog.wait_window()
        return result["value"]

    def _start(self):
        project = self.selected_project.get().strip()
        if project not in project_names(self.projects):
            messagebox.showwarning("Projet manquant", "Sélectionnez un projet valide dans la liste.")
            return
        self.tracking   = True
        self.start_time = datetime.now()
        self.elapsed    = 0
        self._running   = True
        self._timer_gen += 1
        tick_gen = self._timer_gen
        set_button_palette(
            self.start_btn,
            bg=THEME["surface"],
            fg=THEME["muted"],
            hover_bg=THEME["surface"],
            state="disabled",
        )
        set_button_palette(
            self.stop_btn,
            bg=THEME["danger"],
            fg=THEME["white"],
            hover_bg=THEME["danger_hover"],
            state="normal",
        )
        self.status_label.config(
            text=f"● En cours — {self.selected_project.get()}",
            fg=THEME["green"],
        )
        threading.Thread(target=self._tick, args=(tick_gen,), daemon=True).start()

    def _apply_timer_text(self, tick_gen: int, text: str) -> None:
        if tick_gen != self._timer_gen:
            return
        self.timer_label.config(text=text)

    def _tick(self, tick_gen: int) -> None:
        while self._running and tick_gen == self._timer_gen:
            self.elapsed = (datetime.now() - self.start_time).total_seconds()
            h = int(self.elapsed // 3600)
            m = int((self.elapsed % 3600) // 60)
            s = int(self.elapsed % 60)
            text = f"{h:02}:{m:02}:{s:02}"
            self.root.after(0, self._apply_timer_text, tick_gen, text)
            time.sleep(1)

    def _stop(self):
        if not self.tracking:
            return
        self._running = False
        self.tracking = False
        self._timer_gen += 1
        end_time = datetime.now()
        project  = self.selected_project.get()
        start_time = self.start_time
        elapsed = self.elapsed

        self._set_busy_controls()
        self.status_label.config(text="● Enregistrement…", fg=THEME["muted"])

        notes = self._ask_session_notes()
        duration_label = format_session_duration(elapsed)

        try:
            if start_time is None:
                raise ExcelWriteError("Horodatage de début manquant.")
            log_session(project, start_time, end_time, notes, self.excel_file)
        except ExcelWriteError as exc:
            self._set_idle_controls()
            self.timer_label.config(text="00:00:00")
            self.status_label.config(
                text="⚠ Non enregistré",
                fg=THEME["danger"],
            )
            messagebox.showerror(
                "Enregistrement impossible",
                f"La session n'a pas été écrite dans Excel.\n\n{exc}",
                parent=self.root,
            )
            if "vrai .xlsx" in str(exc) or "illisible ou corrompu" in str(exc):
                self._offer_recreate_excel()
            self.status_label.config(text="● En attente", fg=THEME["muted"])
            return

        self._set_idle_controls()
        self.timer_label.config(text="00:00:00")
        self.status_label.config(text="✔ Session enregistrée", fg=THEME["green_dark"])
        messagebox.showinfo(
            "Session enregistrée",
            f"✔ {project}\nDurée : {duration_label}\n\nEnregistré dans :\n{self.excel_file}\n\n"
            "Si Calc/Excel était ouvert : fermez puis rouvrez le fichier pour voir la ligne.",
            parent=self.root,
        )
        self.status_label.config(text="● En attente", fg=THEME["muted"])

    def _offer_recreate_excel(self) -> None:
        if not messagebox.askyesno(
            "Recréer le fichier Excel ?",
            "Créer un nouveau fichier .xlsx valide à cet emplacement ?\n\n"
            "L'ancien fichier sera renommé avec l'extension .bak "
            "(vos anciennes lignes Calc restent dans le .bak).",
            parent=self.root,
        ):
            return
        excel_path = self.excel_file
        if os.path.exists(excel_path):
            bak_path = f"{excel_path}.bak"
            suffix = 1
            while os.path.exists(bak_path):
                bak_path = f"{excel_path}.bak.{suffix}"
                suffix += 1
            try:
                os.replace(excel_path, bak_path)
            except OSError as exc:
                messagebox.showerror(
                    "Sauvegarde impossible",
                    f"Impossible de renommer l'ancien fichier :\n{exc}",
                    parent=self.root,
                )
                return
        try:
            ensure_excel(excel_path)
        except ExcelWriteError as exc:
            messagebox.showerror(
                "Création impossible",
                str(exc),
                parent=self.root,
            )
            return
        messagebox.showinfo(
            "Fichier prêt",
            f"Nouveau fichier créé :\n{excel_path}\n\n"
            "Relancez Start → Stop pour enregistrer la session.",
            parent=self.root,
        )

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