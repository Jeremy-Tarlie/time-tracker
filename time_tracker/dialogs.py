from __future__ import annotations

import os
import tkinter as tk
from collections.abc import Callable
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import cast

from time_tracker.config import project_names, save_config, sort_projects
from time_tracker.constants import (
    APP_NAME,
    APP_VERSION,
    DEFAULT_SORT,
    SORT_LABELS,
    SORT_OPTIONS,
    HotkeyAction,
    HotkeysConfig,
    ProjectEntry,
    SessionRow,
    SortMode,
)
from time_tracker.excel import (
    ExcelWriteError,
    delete_session_row,
    format_session_duration,
    hours_by_project,
    list_sessions_for_day,
    list_sessions_in_range,
    month_bounds,
    open_excel_file,
    parse_session_datetimes,
    rename_project_in_excel,
    total_session_hours,
    update_session_row,
    week_bounds,
)
from time_tracker.hotkeys import (
    HOTKEY_ACTIONS,
    HOTKEY_LABELS,
    combo_from_tk_event,
    default_hotkeys,
    find_hotkey_conflicts,
    format_hotkey_display,
    format_hotkeys_help,
)
from time_tracker.paths import APP_DIR, default_excel_path, normalize_excel_path
from time_tracker.theme import (
    THEME,
    apply_ttk_theme,
    apply_window,
    dialog_header,
    font_ui,
    make_button,
    stat_pill,
    style_checkbutton,
    style_entry,
    style_session_listbox,
    surface_card,
)
from time_tracker.widgets import SearchableDropdown

_ALL_PROJECTS = "Tous"
_DATE_FMT = "%d/%m/%Y"


class SegmentedPresets(tk.Frame):
    """Barre de presets avec état actif (accent)."""

    def __init__(
        self,
        parent: tk.Misc,
        options: list[tuple[str, Callable[[], None]]],
        *,
        initial: str,
    ) -> None:
        t = THEME
        super().__init__(
            parent,
            bg=t.field,
            highlightthickness=1,
            highlightbackground=t.border,
            padx=3,
            pady=3,
        )
        self._buttons: dict[str, tk.Button] = {}
        self._active = initial
        for key, command in options:

            def _make_cmd(k: str = key, c: Callable[[], None] = command) -> Callable[[], None]:
                def run() -> None:
                    self.set_active(k)
                    c()

                return run

            btn = tk.Button(
                self,
                text=key,
                command=_make_cmd(),
                bd=0,
                highlightthickness=0,
                relief="flat",
                cursor="hand2",
                padx=12,
                pady=7,
                font=font_ui(9, bold=False),
            )
            btn.pack(side="left", padx=1)
            self._buttons[key] = btn
        self.set_active(initial)

    def set_active(self, key: str) -> None:
        self._active = key
        t = THEME
        for name, btn in self._buttons.items():
            if name == key:
                btn.configure(
                    bg=t.accent,
                    fg=t.on_accent,
                    activebackground=t.accent_dim,
                    activeforeground=t.on_accent,
                    font=font_ui(9, bold=True),
                )
            else:
                btn.configure(
                    bg=t.field,
                    fg=t.muted,
                    activebackground=t.border,
                    activeforeground=t.text,
                    font=font_ui(9, bold=False),
                )

    def clear_active(self) -> None:
        self._active = ""
        t = THEME
        for btn in self._buttons.values():
            btn.configure(
                bg=t.field,
                fg=t.muted,
                activebackground=t.border,
                activeforeground=t.text,
                font=font_ui(9, bold=False),
            )


def _format_session_line(row: SessionRow, *, with_date: bool = False) -> str:
    notes = f"  ·  {row['notes']}" if row["notes"] else ""
    date_part = f"{row['date']}  ·  " if with_date else ""
    return (
        f"{date_part}{row['start']}–{row['end']}  ·  {row['project']}  ·  {row['duration']}{notes}"
    )


def _session_actions_row(
    parent: tk.Misc,
    *,
    on_edit: Callable[[], None],
    on_delete: Callable[[], None],
    on_close: Callable[[], None],
) -> tk.Frame:
    t = THEME
    btn_row = tk.Frame(parent, bg=t.bg)
    make_button(
        btn_row,
        "Modifier…",
        on_edit,
        variant="subtle",
        font_size=9,
        bold=False,
    ).pack(side="left")
    make_button(
        btn_row,
        "Supprimer",
        on_delete,
        variant="danger",
        font_size=9,
        bold=False,
    ).pack(side="left", padx=(8, 0))
    make_button(
        btn_row,
        "Fermer",
        on_close,
        variant="subtle",
        font_size=9,
        bold=False,
    ).pack(side="right")
    return btn_row


def _clear_pills(frame: tk.Frame) -> None:
    for child in frame.winfo_children():
        child.destroy()


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
        apply_window(self)
        self._build()
        self.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        self.geometry(f"500x220+{px - 250}+{py}")

    def _build(self) -> None:
        t = THEME
        dialog_header(self, "Fichier Excel")
        body = tk.Frame(self, bg=t.bg, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        tk.Label(
            body,
            text="Emplacement et nom du fichier d'export :",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(anchor="w", pady=(0, 6))

        path_row = tk.Frame(body, bg=t.bg)
        path_row.pack(fill="x", pady=(0, 8))
        self._path_entry = tk.Entry(
            path_row,
            textvariable=self._path_var,
            font=font_ui(9),
        )
        style_entry(self._path_entry)
        self._path_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))
        make_button(
            path_row,
            "Ouvrir",
            self._open,
            variant="subtle",
            font_size=9,
            bold=False,
        ).pack(side="left", padx=(0, 6))
        make_button(
            path_row,
            "Parcourir…",
            self._browse,
            variant="primary",
            font_size=9,
        ).pack(side="left")

        tk.Label(
            body,
            text="Les nouvelles sessions seront enregistrées à cet emplacement.\n"
            "L'ancien fichier n'est pas déplacé automatiquement.",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8),
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        btn_row = tk.Frame(body, bg=t.bg)
        btn_row.pack(fill="x")
        make_button(
            btn_row,
            "Par défaut",
            self._use_default,
            variant="subtle",
            font_size=9,
            bold=False,
        ).pack(side="left")
        make_button(
            btn_row,
            "Annuler",
            self.destroy,
            variant="subtle",
            font_size=9,
            bold=False,
        ).pack(side="right", padx=(6, 0))
        make_button(
            btn_row,
            "Enregistrer",
            self._save,
            variant="primary",
            font_size=9,
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
                messagebox.showerror(
                    "Erreur",
                    f"Impossible de créer le dossier :\n{exc}",
                    parent=self,
                )
                return
        self._on_save(excel_path)
        self.destroy()


class ManageProjectsDialog(tk.Toplevel):
    """Fenêtre modale pour gérer (modifier / supprimer) les projets."""

    def __init__(
        self,
        parent: tk.Misc,
        projects: list[ProjectEntry],
        sort_mode: SortMode,
        excel_file: str,
        on_close: Callable[[list[ProjectEntry], SortMode], None],
        *,
        launch_at_startup: bool,
        always_on_top: bool,
        window_geometry: str | None,
        hotkeys: HotkeysConfig,
    ) -> None:
        super().__init__(parent)
        self.projects: list[ProjectEntry] = list(projects)
        self.excel_file = excel_file
        self.launch_at_startup = launch_at_startup
        self.always_on_top = always_on_top
        self.window_geometry = window_geometry
        self.hotkeys = hotkeys
        self.sort_mode = tk.StringVar(value=SORT_LABELS[sort_mode])
        names = project_names(sort_projects(self.projects, sort_mode))
        self.selected_name = tk.StringVar(value=names[0] if names else "")
        self.on_close = on_close
        self.title("Gérer les projets")
        self.resizable(True, True)
        self.minsize(320, 320)
        self.grab_set()
        apply_window(self)
        apply_ttk_theme()
        self._build()
        self._refresh_picker()
        self.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        self.geometry(f"340x300+{px - 260}+{py}")

    def _build(self) -> None:
        t = THEME
        dialog_header(self, "Gérer les projets")
        body = tk.Frame(self, bg=t.bg, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        top_row = tk.Frame(body, bg=t.bg)
        top_row.pack(fill="x", pady=(0, 6))
        tk.Label(
            top_row,
            text="Projet",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(side="left")

        sort_combo = ttk.Combobox(
            top_row,
            textvariable=self.sort_mode,
            state="readonly",
            values=[label for label, _mode in SORT_OPTIONS],
            font=font_ui(8),
            width=8,
            style="App.TCombobox",
        )
        sort_combo.pack(side="right", padx=(8, 0))
        sort_combo.bind("<<ComboboxSelected>>", self._on_sort_selected)

        self.project_picker = SearchableDropdown(
            body,
            self.selected_name,
            bg=t.bg,
            field_bg=t.surface,
            border=t.border,
            muted=t.muted,
            accent=t.accent_dim,
            highlight=t.accent,
            text=t.text,
            placeholder="Choisir un projet…",
        )
        self.project_picker.pack(fill="x", pady=(0, 12))
        self.bind("<Configure>", self._on_dialog_configure)

        btn_frame = tk.Frame(body, bg=t.bg)
        btn_frame.pack(fill="x")
        make_button(btn_frame, "Renommer", self._rename, variant="primary", font_size=9).pack(
            side="left",
            padx=(0, 6),
        )
        make_button(btn_frame, "Supprimer", self._delete, variant="danger", font_size=9).pack(
            side="left",
        )

        tk.Frame(body, bg=t.border, height=1).pack(fill="x", pady=(12, 8))
        tk.Label(
            body,
            text="Ajouter un projet",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(anchor="w", pady=(0, 4))
        add_row = tk.Frame(body, bg=t.bg)
        add_row.pack(fill="x")
        self.new_var = tk.StringVar()
        entry = tk.Entry(add_row, textvariable=self.new_var, font=font_ui(10))
        style_entry(entry)
        entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))
        entry.bind("<Return>", lambda _event: self._add())
        make_button(add_row, "Ajouter", self._add, variant="primary", font_size=9).pack(side="left")
        make_button(
            body,
            "Fermer",
            self._close,
            variant="subtle",
            font_size=9,
            bold=False,
        ).pack(anchor="e", pady=(12, 0))

    def _current_sort_mode(self) -> SortMode:
        label = self.sort_mode.get()
        for option_label, mode in SORT_OPTIONS:
            if option_label == label:
                return mode
        return DEFAULT_SORT

    def _on_dialog_configure(self, event: tk.Event[tk.Misc]) -> None:
        if event.widget is self:
            self.project_picker.close_if_open()

    def _on_sort_selected(self, _event: object = None) -> None:
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
            messagebox.showinfo(
                "Sélection",
                "Sélectionnez d'abord un projet.",
                parent=self,
            )
            return None
        return name

    def _find_project(self, name: str) -> ProjectEntry | None:
        for entry in self.projects:
            if entry["name"] == name:
                return entry
        return None

    def _rename(self) -> None:
        old_name = self._require_selection()
        if old_name is None:
            return
        old_entry = self._find_project(old_name)
        if old_entry is None:
            return
        new = simpledialog.askstring(
            "Renommer",
            f"Nouveau nom pour « {old_name} » :",
            initialvalue=old_name,
            parent=self,
        )
        if not new or not new.strip() or new.strip() == old_name:
            return
        new_name = new.strip()
        if any(entry["name"] == new_name for entry in self.projects):
            messagebox.showwarning("Doublon", f"« {new_name} » existe déjà.", parent=self)
            return
        try:
            rename_project_in_excel(old_name, new_name, self.excel_file)
        except ExcelWriteError as exc:
            messagebox.showerror(
                "Renommage Excel impossible",
                f"Le projet n'a pas été renommé dans Excel.\n"
                f"La liste locale n'a pas été modifiée.\n\n{exc}",
                parent=self,
            )
            return
        old_entry["name"] = new_name
        self.selected_name.set(new_name)
        self._refresh_picker()

    def _delete(self) -> None:
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

    def _add(self) -> None:
        name = self.new_var.get().strip()
        if not name:
            return
        if any(entry["name"] == name for entry in self.projects):
            messagebox.showwarning("Doublon", f"« {name} » existe déjà.", parent=self)
            return
        self.projects.append(
            {
                "name": name,
                "created": datetime.now().isoformat(timespec="seconds"),
            }
        )
        self.new_var.set("")
        self.selected_name.set(name)
        self._refresh_picker()

    def _close(self) -> None:
        self.destroy()

    def destroy(self) -> None:
        self.project_picker.close_if_open()
        save_config(
            {
                "projects": self.projects,
                "sort_mode": self._current_sort_mode(),
                "excel_file": self.excel_file,
                "launch_at_startup": self.launch_at_startup,
                "always_on_top": self.always_on_top,
                "window_geometry": self.window_geometry,
                "hotkeys": self.hotkeys,
            }
        )
        self.on_close(self.projects, self._current_sort_mode())
        self.grab_release()
        super().destroy()


class SettingsDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        launch_at_startup: bool,
        always_on_top: bool,
        hotkeys: HotkeysConfig,
        on_save: Callable[[bool, bool, HotkeysConfig], None],
        startup_available: bool = True,
        startup_label: str = "Lancer au démarrage",
    ) -> None:
        super().__init__(parent)
        self._on_save = on_save
        self._startup_var = tk.BooleanVar(value=launch_at_startup)
        self._topmost_var = tk.BooleanVar(value=always_on_top)
        self._hotkeys: HotkeysConfig = {
            "toggle": hotkeys["toggle"],
            "stop": hotkeys["stop"],
            "settings": hotkeys["settings"],
            "show": hotkeys["show"],
            "today": hotkeys["today"],
            "period": hotkeys["period"],
            "manual": hotkeys["manual"],
        }
        self._capture_action: HotkeyAction | None = None
        self._hotkey_vars: dict[HotkeyAction, tk.StringVar] = {}
        self.title("Réglages")
        apply_window(self)
        self.resizable(False, False)
        self.grab_set()
        t = THEME
        dialog_header(self, "Réglages")
        body = tk.Frame(self, bg=t.bg, padx=18, pady=16)
        body.pack(fill="both", expand=True)

        tk.Label(
            body,
            text="Fenêtre",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(anchor="w", pady=(0, 6))
        topmost_cb = tk.Checkbutton(
            body,
            text="Rester toujours au premier plan",
            variable=self._topmost_var,
        )
        style_checkbutton(topmost_cb)
        topmost_cb.pack(anchor="w")
        tk.Label(
            body,
            text="Si désactivé, la fenêtre peut passer derrière les autres apps.",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8),
            wraplength=360,
            justify="left",
        ).pack(anchor="w", pady=(2, 14))

        tk.Label(
            body,
            text="Démarrage",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(anchor="w", pady=(0, 6))
        cb = tk.Checkbutton(
            body,
            text=startup_label,
            variable=self._startup_var,
            state="normal" if startup_available else "disabled",
        )
        style_checkbutton(cb)
        cb.pack(anchor="w")
        if not startup_available:
            tk.Label(
                body,
                text="Indisponible en mode portable ou sur cet OS.",
                bg=t.bg,
                fg=t.muted,
                font=font_ui(8),
            ).pack(anchor="w", pady=(4, 0))

        tk.Label(
            body,
            text="Raccourcis globaux",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(anchor="w", pady=(14, 6))
        tk.Label(
            body,
            text="Cliquez « Modifier », puis tapez la combinaison.",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8),
        ).pack(anchor="w", pady=(0, 6))

        for action in HOTKEY_ACTIONS:
            row = tk.Frame(body, bg=t.bg)
            row.pack(fill="x", pady=2)
            tk.Label(
                row,
                text=HOTKEY_LABELS[action],
                bg=t.bg,
                fg=t.text,
                font=font_ui(8),
                width=18,
                anchor="w",
            ).pack(side="left")
            var = tk.StringVar(value=format_hotkey_display(self._hotkeys[action]))
            self._hotkey_vars[action] = var
            entry = tk.Entry(
                row,
                textvariable=var,
                font=font_ui(8),
                width=16,
                state="readonly",
                readonlybackground=t.field,
                fg=t.text,
                relief="flat",
                highlightthickness=1,
                highlightbackground=t.border,
            )
            entry.pack(side="left", padx=(0, 6), ipady=3)

            def _start(act: HotkeyAction = action) -> None:
                self._begin_capture(act)

            make_button(
                row,
                "Modifier",
                _start,
                variant="subtle",
                font_size=8,
                bold=False,
            ).pack(side="left")

        make_button(
            body,
            "Réinitialiser les raccourcis",
            self._reset_hotkeys,
            variant="link",
            font_size=8,
            bold=False,
        ).pack(anchor="w", pady=(8, 0))

        self.bind("<KeyPress>", self._on_capture_key)

        btn_row = tk.Frame(body, bg=t.bg)
        btn_row.pack(fill="x", pady=(18, 0))
        make_button(
            btn_row,
            "À propos",
            self._open_about,
            variant="link",
            font_size=9,
            bold=False,
        ).pack(side="left")
        make_button(
            btn_row,
            "Annuler",
            self.destroy,
            variant="subtle",
            font_size=9,
            bold=False,
        ).pack(side="right", padx=(6, 0))
        make_button(btn_row, "Enregistrer", self._save, variant="primary", font_size=9).pack(
            side="right",
        )
        self.minsize(420, 560)
        self.geometry("420x600")

    def _refresh_hotkey_labels(self) -> None:
        for action in HOTKEY_ACTIONS:
            self._hotkey_vars[action].set(format_hotkey_display(self._hotkeys[action]))

    def _reset_hotkeys(self) -> None:
        self._capture_action = None
        self._hotkeys = default_hotkeys()
        self._refresh_hotkey_labels()

    def _begin_capture(self, action: HotkeyAction) -> None:
        self._capture_action = action
        self._hotkey_vars[action].set("…")
        self.focus_force()

    def _on_capture_key(self, event: tk.Event[tk.Misc]) -> str | None:
        if self._capture_action is None:
            return None
        if event.keysym.lower() == "escape":
            action = self._capture_action
            self._capture_action = None
            self._hotkey_vars[action].set(format_hotkey_display(self._hotkeys[action]))
            return "break"
        combo = combo_from_tk_event(event.keysym, int(event.state))
        if combo is None:
            return "break"
        action = self._capture_action
        self._capture_action = None
        self._hotkeys[action] = combo
        self._hotkey_vars[action].set(format_hotkey_display(combo))
        return "break"

    def _open_about(self) -> None:
        AboutDialog(self, hotkeys=self._hotkeys)

    def _save(self) -> None:
        conflicts = find_hotkey_conflicts(self._hotkeys)
        if conflicts:
            messagebox.showwarning(
                "Raccourcis",
                "Deux actions partagent le même raccourci :\n"
                + "\n".join(format_hotkey_display(c) for c in conflicts),
                parent=self,
            )
            return
        saved: HotkeysConfig = {
            "toggle": self._hotkeys["toggle"],
            "stop": self._hotkeys["stop"],
            "settings": self._hotkeys["settings"],
            "show": self._hotkeys["show"],
            "today": self._hotkeys["today"],
            "period": self._hotkeys["period"],
            "manual": self._hotkeys["manual"],
        }
        self._on_save(
            bool(self._startup_var.get()),
            bool(self._topmost_var.get()),
            saved,
        )
        self.destroy()


class AboutDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        hotkeys: HotkeysConfig | None = None,
    ) -> None:
        super().__init__(parent)
        self.title("À propos")
        apply_window(self)
        self.resizable(False, False)
        self.grab_set()
        t = THEME
        dialog_header(self, "À propos")
        body = tk.Frame(self, bg=t.bg, padx=18, pady=16)
        body.pack(fill="both", expand=True)
        tk.Label(
            body,
            text=f"{APP_NAME} {APP_VERSION}",
            bg=t.bg,
            fg=t.text,
            font=font_ui(12, bold=True),
        ).pack(anchor="w")
        tk.Label(
            body,
            text=f"Dossier données :\n{APP_DIR}",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8),
            justify="left",
            wraplength=360,
        ).pack(anchor="w", pady=(10, 0))
        log_path = os.path.join(APP_DIR, "timetracker.log")
        tk.Label(
            body,
            text=f"Journal :\n{log_path}",
            bg=t.bg,
            fg=t.muted,
            font=font_ui(8),
            justify="left",
            wraplength=360,
        ).pack(anchor="w", pady=(8, 0))
        help_text = format_hotkeys_help(hotkeys or default_hotkeys())
        tk.Label(
            body,
            text="Raccourcis\n" + help_text,
            bg=t.bg,
            fg=t.text,
            font=font_ui(8),
            justify="left",
        ).pack(anchor="w", pady=(14, 0))
        make_button(
            body,
            "Fermer",
            self.destroy,
            variant="subtle",
            font_size=9,
            bold=False,
        ).pack(anchor="e", pady=(16, 0))
        self.geometry("420x420")


class EditSessionDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        projects: list[ProjectEntry],
        session: SessionRow,
        on_saved: Callable[[], None],
        excel_file: str,
    ) -> None:
        super().__init__(parent)
        self._projects = projects
        self._session = session
        self._on_saved = on_saved
        self._excel_file = excel_file
        names = project_names(projects)
        self._project = tk.StringVar(value=session["project"])
        try:
            start_dt, end_dt = parse_session_datetimes(session)
        except ValueError:
            start_dt = datetime.now().replace(second=0, microsecond=0)
            end_dt = start_dt + timedelta(hours=1)
        self._start_var = tk.StringVar(value=start_dt.strftime("%Y-%m-%d %H:%M"))
        self._end_var = tk.StringVar(value=end_dt.strftime("%Y-%m-%d %H:%M"))
        self._notes_var = tk.StringVar(value=session["notes"])
        self.title("Modifier la session")
        apply_window(self)
        self.resizable(False, False)
        self.grab_set()
        t = THEME
        dialog_header(self, "Modifier la session")
        body = tk.Frame(self, bg=t.bg, padx=16, pady=14)
        body.pack(fill="both", expand=True)
        self._picker = SearchableDropdown(
            body,
            self._project,
            bg=t.bg,
            field_bg=t.surface,
            border=t.border,
            muted=t.muted,
            accent=t.accent_dim,
            highlight=t.accent,
            text=t.text,
        )
        self._picker.set_options(names)
        self._picker.pack(fill="x", pady=(0, 8))
        for label, var in (
            ("Début (AAAA-MM-JJ HH:MM)", self._start_var),
            ("Fin (AAAA-MM-JJ HH:MM)", self._end_var),
            ("Notes", self._notes_var),
        ):
            tk.Label(body, text=label, bg=t.bg, fg=t.muted, font=font_ui(8, bold=True)).pack(
                anchor="w",
            )
            entry = tk.Entry(body, textvariable=var, font=font_ui(10))
            style_entry(entry)
            entry.pack(fill="x", pady=(2, 8), ipady=4)
        btn_row = tk.Frame(body, bg=t.bg)
        btn_row.pack(fill="x", pady=(4, 0))
        make_button(
            btn_row,
            "Annuler",
            self.destroy,
            variant="subtle",
            font_size=9,
            bold=False,
        ).pack(side="right", padx=(6, 0))
        make_button(btn_row, "Enregistrer", self._submit, variant="primary", font_size=9).pack(
            side="right",
        )
        self.geometry("360x340")

    def _submit(self) -> None:
        project = self._project.get().strip()
        if project not in project_names(self._projects):
            messagebox.showwarning("Projet", "Sélectionnez un projet valide.", parent=self)
            return
        try:
            start = datetime.strptime(self._start_var.get().strip(), "%Y-%m-%d %H:%M")
            end = datetime.strptime(self._end_var.get().strip(), "%Y-%m-%d %H:%M")
        except ValueError:
            messagebox.showwarning(
                "Format",
                "Utilisez le format AAAA-MM-JJ HH:MM pour début et fin.",
                parent=self,
            )
            return
        if end <= start:
            messagebox.showwarning("Horaires", "La fin doit être après le début.", parent=self)
            return
        try:
            update_session_row(
                self._excel_file,
                self._session["row_index"],
                project=project,
                start_dt=start,
                end_dt=end,
                notes=self._notes_var.get().strip(),
            )
        except ExcelWriteError as exc:
            messagebox.showerror("Modification impossible", str(exc), parent=self)
            return
        self._on_saved()
        self.destroy()


class TodayHistoryDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        excel_file: str,
        projects: list[ProjectEntry],
    ) -> None:
        super().__init__(parent)
        self._excel_file = excel_file
        self._projects = projects
        self._rows: list[SessionRow] = []
        self.title("Aujourd'hui")
        apply_window(self)
        self.minsize(520, 420)
        self.grab_set()
        t = THEME
        dialog_header(self, "Sessions du jour")
        body = tk.Frame(self, bg=t.bg, padx=16, pady=14)
        body.pack(fill="both", expand=True)

        self._pills = tk.Frame(body, bg=t.bg)
        self._pills.pack(fill="x", pady=(0, 12))

        list_outer, list_inner = surface_card(body, padx=10, pady=10)
        list_outer.pack(fill="both", expand=True)
        tk.Label(
            list_inner,
            text="SESSIONS",
            bg=t.surface,
            fg=t.muted,
            font=font_ui(8, bold=True),
            anchor="w",
        ).pack(fill="x", pady=(0, 8))
        self._listbox = tk.Listbox(list_inner, height=12)
        style_session_listbox(self._listbox)
        self._listbox.pack(fill="both", expand=True)

        _session_actions_row(
            body,
            on_edit=self._edit_selected,
            on_delete=self._delete_selected,
            on_close=self.destroy,
        ).pack(fill="x", pady=(12, 0))
        self.geometry("540x480")
        self._reload()

    def _reload(self) -> None:
        self._rows = list_sessions_for_day(self._excel_file)
        hours = total_session_hours(self._rows)
        _clear_pills(self._pills)
        stat_pill(self._pills, "Sessions", str(len(self._rows))).pack(side="left", padx=(0, 8))
        stat_pill(
            self._pills,
            "Total",
            format_session_duration(hours * 3600),
            accent=True,
        ).pack(side="left")
        self._listbox.delete(0, "end")
        if not self._rows:
            self._listbox.insert("end", "Aucune session enregistrée aujourd'hui.")
            return
        for row in self._rows:
            self._listbox.insert("end", _format_session_line(row))

    def _selected_session(self) -> SessionRow | None:
        if not self._rows:
            return None
        selection = cast(tuple[int, ...], self._listbox.curselection())
        if not selection:
            messagebox.showwarning("Sélection", "Choisissez une session.", parent=self)
            return None
        index = selection[0]
        if index < 0 or index >= len(self._rows):
            return None
        return self._rows[index]

    def _edit_selected(self) -> None:
        session = self._selected_session()
        if session is None:
            return
        if not self._projects:
            messagebox.showwarning("Projets", "Ajoutez un projet avant de modifier.", parent=self)
            return
        EditSessionDialog(self, self._projects, session, self._reload, self._excel_file)

    def _delete_selected(self) -> None:
        session = self._selected_session()
        if session is None:
            return
        if not messagebox.askyesno(
            "Supprimer",
            f"Supprimer la session {session['start']}–{session['end']} ({session['project']}) ?",
            parent=self,
        ):
            return
        try:
            delete_session_row(self._excel_file, session["row_index"])
        except ExcelWriteError as exc:
            messagebox.showerror("Suppression impossible", str(exc), parent=self)
            return
        self._reload()


class PeriodHistoryDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        excel_file: str,
        projects: list[ProjectEntry],
    ) -> None:
        super().__init__(parent)
        self._excel_file = excel_file
        self._projects = projects
        self._rows: list[SessionRow] = []
        self._start_day, self._end_day = week_bounds()
        self._preset_key = "Cette semaine"
        self.title("Période")
        apply_window(self)
        apply_ttk_theme()
        self.minsize(600, 560)
        self.grab_set()
        t = THEME
        dialog_header(self, "Sessions par période")
        body = tk.Frame(self, bg=t.bg, padx=16, pady=14)
        body.pack(fill="both", expand=True)

        self._presets = SegmentedPresets(
            body,
            [
                ("Aujourd'hui", self._preset_today),
                ("Cette semaine", self._preset_week),
                ("Ce mois", self._preset_month),
            ],
            initial="Cette semaine",
        )
        self._presets.pack(anchor="w", pady=(0, 12))

        filters_outer, filters_inner = surface_card(body, padx=14, pady=12)
        filters_outer.pack(fill="x", pady=(0, 12))
        tk.Label(
            filters_inner,
            text="FILTRES",
            bg=t.surface,
            fg=t.muted,
            font=font_ui(8, bold=True),
            anchor="w",
        ).pack(fill="x", pady=(0, 10))

        range_row = tk.Frame(filters_inner, bg=t.surface)
        range_row.pack(fill="x", pady=(0, 10))
        tk.Label(
            range_row,
            text="Du",
            bg=t.surface,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(side="left")
        self._start_var = tk.StringVar(value=self._start_day.strftime(_DATE_FMT))
        start_entry = tk.Entry(
            range_row,
            textvariable=self._start_var,
            width=12,
            font=font_ui(9),
        )
        style_entry(start_entry)
        start_entry.pack(side="left", padx=(8, 12), ipady=4)
        tk.Label(
            range_row,
            text="au",
            bg=t.surface,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(side="left")
        self._end_var = tk.StringVar(value=self._end_day.strftime(_DATE_FMT))
        end_entry = tk.Entry(range_row, textvariable=self._end_var, width=12, font=font_ui(9))
        style_entry(end_entry)
        end_entry.pack(side="left", padx=(8, 12), ipady=4)
        make_button(
            range_row,
            "Appliquer",
            self._apply_custom_range,
            variant="primary",
            font_size=9,
        ).pack(side="left")

        filter_row = tk.Frame(filters_inner, bg=t.surface)
        filter_row.pack(fill="x")
        tk.Label(
            filter_row,
            text="Projet",
            bg=t.surface,
            fg=t.muted,
            font=font_ui(8, bold=True),
        ).pack(side="left")
        project_names_list = [_ALL_PROJECTS, *project_names(projects)]
        self._project_filter = tk.StringVar(value=_ALL_PROJECTS)
        project_combo = ttk.Combobox(
            filter_row,
            textvariable=self._project_filter,
            state="readonly",
            values=project_names_list,
            font=font_ui(9),
            width=24,
            style="App.TCombobox",
        )
        project_combo.pack(side="left", padx=(8, 0))
        project_combo.bind("<<ComboboxSelected>>", lambda _event: self._reload())

        self._pills = tk.Frame(body, bg=t.bg)
        self._pills.pack(fill="x", pady=(0, 8))
        self._by_project = tk.Frame(body, bg=t.bg)
        self._by_project.pack(fill="x", pady=(0, 12))

        list_outer, list_inner = surface_card(body, padx=10, pady=10)
        list_outer.pack(fill="both", expand=True)
        tk.Label(
            list_inner,
            text="SESSIONS",
            bg=t.surface,
            fg=t.muted,
            font=font_ui(8, bold=True),
            anchor="w",
        ).pack(fill="x", pady=(0, 8))
        self._listbox = tk.Listbox(list_inner, height=12)
        style_session_listbox(self._listbox)
        self._listbox.pack(fill="both", expand=True)

        _session_actions_row(
            body,
            on_edit=self._edit_selected,
            on_delete=self._delete_selected,
            on_close=self.destroy,
        ).pack(fill="x", pady=(12, 0))
        self.geometry("620x640")
        self._reload()

    def _preset_today(self) -> None:
        self._preset_key = "Aujourd'hui"
        today = date.today()
        self._set_bounds(today, today, sync_preset=False)

    def _preset_week(self) -> None:
        self._preset_key = "Cette semaine"
        self._set_bounds(*week_bounds(), sync_preset=False)

    def _preset_month(self) -> None:
        self._preset_key = "Ce mois"
        self._set_bounds(*month_bounds(), sync_preset=False)

    def _set_bounds(
        self,
        start_day: date,
        end_day: date,
        *,
        sync_preset: bool = True,
    ) -> None:
        self._start_day = start_day
        self._end_day = end_day
        self._start_var.set(start_day.strftime(_DATE_FMT))
        self._end_var.set(end_day.strftime(_DATE_FMT))
        if sync_preset:
            self._preset_key = "Perso"
            self._presets.clear_active()
        self._reload()

    def _parse_date_field(self, raw: str) -> date | None:
        try:
            return datetime.strptime(raw.strip(), _DATE_FMT).date()
        except ValueError:
            return None

    def _apply_custom_range(self) -> None:
        start_day = self._parse_date_field(self._start_var.get())
        end_day = self._parse_date_field(self._end_var.get())
        if start_day is None or end_day is None:
            messagebox.showwarning(
                "Dates",
                f"Utilisez le format {_DATE_FMT} (ex. 09/11/2026).",
                parent=self,
            )
            return
        if end_day < start_day:
            start_day, end_day = end_day, start_day
        self._set_bounds(start_day, end_day, sync_preset=True)

    def _filtered_rows(self) -> list[SessionRow]:
        rows = list_sessions_in_range(self._excel_file, self._start_day, self._end_day)
        selected = self._project_filter.get().strip()
        if selected and selected != _ALL_PROJECTS:
            rows = [row for row in rows if row["project"] == selected]
        return rows

    def _reload(self) -> None:
        self._rows = self._filtered_rows()
        hours = total_session_hours(self._rows)
        _clear_pills(self._pills)
        range_label = f"{self._start_day.strftime(_DATE_FMT)} → {self._end_day.strftime(_DATE_FMT)}"
        stat_pill(self._pills, "Période", range_label).pack(side="left", padx=(0, 8))
        stat_pill(self._pills, "Sessions", str(len(self._rows))).pack(side="left", padx=(0, 8))
        stat_pill(
            self._pills,
            "Total",
            format_session_duration(hours * 3600),
            accent=True,
        ).pack(side="left")

        _clear_pills(self._by_project)
        by_project = hours_by_project(self._rows)
        if by_project:
            tk.Label(
                self._by_project,
                text="PAR PROJET",
                bg=THEME.bg,
                fg=THEME.muted,
                font=font_ui(8, bold=True),
            ).pack(anchor="w", pady=(0, 6))
            pills_row = tk.Frame(self._by_project, bg=THEME.bg)
            pills_row.pack(fill="x")
            for name, project_hours in by_project:
                stat_pill(
                    pills_row,
                    name,
                    format_session_duration(project_hours * 3600),
                ).pack(side="left", padx=(0, 8), pady=(0, 4))
        self._listbox.delete(0, "end")
        if not self._rows:
            self._listbox.insert("end", "Aucune session pour ces filtres.")
            return
        for row in self._rows:
            self._listbox.insert("end", _format_session_line(row, with_date=True))

    def _selected_session(self) -> SessionRow | None:
        if not self._rows:
            return None
        selection = cast(tuple[int, ...], self._listbox.curselection())
        if not selection:
            messagebox.showwarning("Sélection", "Choisissez une session.", parent=self)
            return None
        index = selection[0]
        if index < 0 or index >= len(self._rows):
            return None
        return self._rows[index]

    def _edit_selected(self) -> None:
        session = self._selected_session()
        if session is None:
            return
        if not self._projects:
            messagebox.showwarning("Projets", "Ajoutez un projet avant de modifier.", parent=self)
            return
        EditSessionDialog(self, self._projects, session, self._reload, self._excel_file)

    def _delete_selected(self) -> None:
        session = self._selected_session()
        if session is None:
            return
        if not messagebox.askyesno(
            "Supprimer",
            f"Supprimer la session {session['start']}–{session['end']} ({session['project']}) ?",
            parent=self,
        ):
            return
        try:
            delete_session_row(self._excel_file, session["row_index"])
        except ExcelWriteError as exc:
            messagebox.showerror("Suppression impossible", str(exc), parent=self)
            return
        self._reload()


class ManualSessionDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        projects: list[ProjectEntry],
        on_submit: Callable[[str, datetime, datetime, str], None],
    ) -> None:
        super().__init__(parent)
        self._on_submit = on_submit
        self._projects = projects
        names = project_names(projects)
        self._project = tk.StringVar(value=names[0] if names else "")
        now = datetime.now().replace(second=0, microsecond=0)
        start = now - timedelta(hours=1)
        self._start_var = tk.StringVar(value=start.strftime("%Y-%m-%d %H:%M"))
        self._end_var = tk.StringVar(value=now.strftime("%Y-%m-%d %H:%M"))
        self._notes_var = tk.StringVar()
        self.title("Saisie manuelle")
        apply_window(self)
        self.resizable(False, False)
        self.grab_set()
        t = THEME
        dialog_header(self, "Ajouter une session")
        body = tk.Frame(self, bg=t.bg, padx=16, pady=14)
        body.pack(fill="both", expand=True)
        self._picker = SearchableDropdown(
            body,
            self._project,
            bg=t.bg,
            field_bg=t.surface,
            border=t.border,
            muted=t.muted,
            accent=t.accent_dim,
            highlight=t.accent,
            text=t.text,
        )
        self._picker.set_options(names)
        self._picker.pack(fill="x", pady=(0, 8))
        for label, var in (
            ("Début (AAAA-MM-JJ HH:MM)", self._start_var),
            ("Fin (AAAA-MM-JJ HH:MM)", self._end_var),
            ("Notes", self._notes_var),
        ):
            tk.Label(body, text=label, bg=t.bg, fg=t.muted, font=font_ui(8, bold=True)).pack(
                anchor="w",
            )
            entry = tk.Entry(body, textvariable=var, font=font_ui(10))
            style_entry(entry)
            entry.pack(fill="x", pady=(2, 8), ipady=4)
        btn_row = tk.Frame(body, bg=t.bg)
        btn_row.pack(fill="x", pady=(4, 0))
        make_button(
            btn_row,
            "Annuler",
            self.destroy,
            variant="subtle",
            font_size=9,
            bold=False,
        ).pack(side="right", padx=(6, 0))
        make_button(btn_row, "Enregistrer", self._submit, variant="primary", font_size=9).pack(
            side="right",
        )
        self.geometry("360x340")

    def _submit(self) -> None:
        project = self._project.get().strip()
        if project not in project_names(self._projects):
            messagebox.showwarning("Projet", "Sélectionnez un projet valide.", parent=self)
            return
        try:
            start = datetime.strptime(self._start_var.get().strip(), "%Y-%m-%d %H:%M")
            end = datetime.strptime(self._end_var.get().strip(), "%Y-%m-%d %H:%M")
        except ValueError:
            messagebox.showwarning(
                "Format",
                "Utilisez le format AAAA-MM-JJ HH:MM pour début et fin.",
                parent=self,
            )
            return
        if end <= start:
            messagebox.showwarning("Horaires", "La fin doit être après le début.", parent=self)
            return
        self._on_submit(project, start, end, self._notes_var.get().strip())
        self.destroy()
