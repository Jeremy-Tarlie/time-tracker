import threading
import time
import tkinter as tk
from collections.abc import Callable
from datetime import datetime
from tkinter import messagebox, simpledialog, ttk

from time_tracker.config import load_config, project_names, save_config, sort_projects
from time_tracker.constants import (
    DEFAULT_SORT,
    SORT_LABELS,
    SORT_OPTIONS,
    ProjectEntry,
    SortMode,
)
from time_tracker.dialogs import ExcelFileDialog, ManageProjectsDialog
from time_tracker.excel import ensure_excel, log_session, open_excel_file
from time_tracker.widgets import SearchableDropdown


class TimeTrackerApp:
    def __init__(self, root: tk.Tk) -> None:
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
        self.tracking = False
        self.start_time: datetime | None = None
        self.elapsed = 0.0
        self._running = False
        self._drag_x = 0
        self._drag_y = 0

        self._build_ui()
        self._refresh_project_list()
        self._position_window()

    def _position_window(self) -> None:
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        self.root.geometry(f"+{sw - w - 30}+{sh - h - 60}")

    def _build_ui(self) -> None:
        bg = "#FFFFFF"
        card = "#FFFFFF"
        green = "#40916C"
        red = "#E63946"
        text = "#1A1A1A"
        muted = "#666666"
        accent = "#2D6A4F"
        subtle = "#F5F5F5"
        border = "#E0E0E0"
        timer_color = "#1B4332"

        self._BG = bg
        self._CARD = card
        self._GREEN = green
        self._RED = red
        self._TEXT = text
        self._MUTED = muted
        self._ACCENT = accent
        self._SUBTLE = subtle
        self._BORDER = border
        self._TIMER = timer_color

        self.root.configure(bg=bg)

        title_bar = tk.Frame(self.root, bg=accent, height=38)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        tk.Label(
            title_bar, text="⏱  Time Tracker", bg=accent, fg="white",
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left", padx=12, pady=7)

        for widget in [title_bar] + title_bar.winfo_children():
            if not isinstance(widget, tk.Button):
                widget.bind("<ButtonPress-1>", self._start_drag)
                widget.bind("<B1-Motion>", self._drag)

        body = tk.Frame(self.root, bg=bg, padx=16, pady=14)
        body.pack(fill="both", expand=True)

        top_row = tk.Frame(body, bg=bg)
        top_row.pack(fill="x", pady=(0, 6))
        tk.Label(
            top_row, text="PROJET", bg=bg, fg=muted, font=("Segoe UI", 7, "bold"),
        ).pack(side="left")

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "W.TCombobox",
            fieldbackground=subtle, background=subtle,
            foreground=text, selectbackground=green,
            selectforeground="white", arrowcolor=accent,
            bordercolor=border, lightcolor=border, darkcolor=border,
        )
        style.map("W.TCombobox", fieldbackground=[("readonly", subtle)])

        sort_combo = ttk.Combobox(
            top_row, textvariable=self.sort_mode, state="readonly",
            values=[label for label, _mode in SORT_OPTIONS],
            font=("Segoe UI", 8), width=8, style="W.TCombobox",
        )
        sort_combo.pack(side="right", padx=(8, 0))
        sort_combo.bind("<<ComboboxSelected>>", self._on_sort_changed)

        tk.Button(
            top_row, text="Gérer ›", bg=bg, fg=green, bd=0,
            font=("Segoe UI", 7, "underline"), cursor="hand2",
            activebackground=bg, activeforeground=accent,
            command=self._open_manage,
        ).pack(side="right")

        self.project_picker = SearchableDropdown(
            body, self.selected_project,
            bg=bg, field_bg=subtle, border=border,
            muted=muted, accent=accent, highlight=green, text=text,
        )
        self.project_picker.pack(fill="x", pady=(0, 12))
        self.root.bind("<Configure>", self._on_root_configure)

        tk.Frame(body, bg=border, height=1).pack(fill="x", pady=(0, 12))

        self.timer_label = tk.Label(
            body, text="00:00:00", bg=bg, fg=timer_color,
            font=("Segoe UI Mono", 32, "bold"),
        )
        self.timer_label.pack()

        self.status_label = tk.Label(
            body, text="● En attente", bg=bg, fg=muted, font=("Segoe UI", 9),
        )
        self.status_label.pack(pady=(2, 14))

        btn_row = tk.Frame(body, bg=bg)
        btn_row.pack(fill="x")

        self.start_btn = tk.Button(
            btn_row, text="▶  Start", bg=green, fg="white",
            font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2",
            padx=14, pady=9, activebackground=accent,
            activeforeground="white", command=self._start,
        )
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.stop_btn = tk.Button(
            btn_row, text="■  Stop", bg=subtle, fg=muted,
            font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2",
            padx=14, pady=9, activebackground=red,
            activeforeground="white", command=self._stop, state="disabled",
        )
        self.stop_btn.pack(side="left", fill="x", expand=True)

        tk.Frame(body, bg=border, height=1).pack(fill="x", pady=(12, 10))

        excel_footer = tk.Frame(body, bg=bg)
        excel_footer.pack(fill="x", pady=(0, 6))

        def excel_link(parent: tk.Frame, label: str, command: Callable[[], None]) -> None:
            tk.Button(
                parent, text=label, bg=bg, fg=muted,
                font=("Segoe UI", 8), bd=0, cursor="hand2",
                activebackground=bg, activeforeground=accent,
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

    def _on_sort_changed(self, _event: object = None) -> None:
        self._refresh_project_list(persist_sort=True)

    def _on_root_configure(self, event: tk.Event[tk.Misc]) -> None:
        if event.widget is self.root:
            self.project_picker.close_if_open()

    def _start_drag(self, event: tk.Event[tk.Misc]) -> None:
        self._drag_x = event.x
        self._drag_y = event.y

    def _drag(self, event: tk.Event[tk.Misc]) -> None:
        x = self.root.winfo_x() + event.x - self._drag_x
        y = self.root.winfo_y() + event.y - self._drag_y
        self.root.geometry(f"+{x}+{y}")

    def _start(self) -> None:
        project = self.selected_project.get().strip()
        if project not in project_names(self.projects):
            messagebox.showwarning(
                "Projet manquant", "Sélectionnez un projet valide dans la liste.",
            )
            return
        self.tracking = True
        self.start_time = datetime.now()
        self.elapsed = 0
        self._running = True
        self.start_btn.config(state="disabled", bg=self._SUBTLE, fg=self._MUTED)
        self.stop_btn.config(state="normal", bg=self._RED, fg="white")
        self.status_label.config(
            text=f"● En cours — {self.selected_project.get()}", fg=self._GREEN,
        )
        threading.Thread(target=self._tick, daemon=True).start()

    def _tick(self) -> None:
        while self._running and self.start_time is not None:
            self.elapsed = (datetime.now() - self.start_time).total_seconds()
            hours = int(self.elapsed // 3600)
            minutes = int((self.elapsed % 3600) // 60)
            seconds = int(self.elapsed % 60)
            self.root.after(
                0,
                self.timer_label.config,
                {"text": f"{hours:02}:{minutes:02}:{seconds:02}"},
            )
            time.sleep(1)

    def _stop(self) -> None:
        if not self.tracking or self.start_time is None:
            return
        self._running = False
        self.tracking = False
        end_time = datetime.now()
        project = self.selected_project.get()
        session_start = self.start_time

        self.start_btn.config(state="normal", bg=self._GREEN, fg="white")
        self.stop_btn.config(state="disabled", bg=self._SUBTLE, fg=self._MUTED)
        self.status_label.config(text="✔ Session enregistrée", fg=self._ACCENT)

        notes = simpledialog.askstring(
            "Notes (optionnel)",
            "Ajouter une note pour cette session :",
            parent=self.root,
        ) or ""

        log_session(project, session_start, end_time, notes, self.excel_file)

        dur = int(self.elapsed)
        hours, minutes = dur // 3600, (dur % 3600) // 60
        messagebox.showinfo(
            "Session enregistrée",
            f"✔ {project}\nDurée : {hours}h {minutes:02}min\n\nEnregistré dans :\n{self.excel_file}",
        )
        self.timer_label.config(text="00:00:00")
        self.status_label.config(text="● En attente", fg=self._MUTED)

    def _open_manage(self) -> None:
        ManageProjectsDialog(
            self.root,
            self.projects,
            self._current_sort_mode(),
            self.excel_file,
            self._on_projects_updated,
        )

    def _on_projects_updated(
        self, new_projects: list[ProjectEntry], sort_mode: SortMode,
    ) -> None:
        self.projects = new_projects
        self.sort_mode.set(SORT_LABELS[sort_mode])
        self._refresh_project_list(persist_sort=True)

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
