import os
import tkinter as tk
from collections.abc import Callable
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk

from time_tracker.config import project_names, save_config, sort_projects
from time_tracker.constants import (
    DEFAULT_SORT,
    SORT_LABELS,
    SORT_OPTIONS,
    ProjectEntry,
    SortMode,
)
from time_tracker.excel import ExcelWriteError, open_excel_file, rename_project_in_excel
from time_tracker.paths import APP_DIR, default_excel_path, normalize_excel_path
from time_tracker.widgets import SearchableDropdown


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
        bg = "#FFFFFF"
        accent = "#40916C"
        green = "#52B788"
        muted = "#888888"
        border = "#E0E0E0"

        hdr = tk.Frame(self, bg=accent, height=36)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="📁  Fichier Excel", bg=accent, fg="white",
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left", padx=12, pady=6)

        body = tk.Frame(self, bg=bg, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        tk.Label(
            body,
            text="Emplacement et nom du fichier d'export :",
            bg=bg, fg=muted, font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", pady=(0, 6))

        path_row = tk.Frame(body, bg=bg)
        path_row.pack(fill="x", pady=(0, 8))
        self._path_entry = tk.Entry(
            path_row, textvariable=self._path_var,
            font=("Segoe UI", 9), bg="#F5F5F5", fg="#222222",
            relief="flat", bd=0, highlightthickness=1,
            highlightbackground=border, highlightcolor=green,
        )
        self._path_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))
        tk.Button(
            path_row, text="Ouvrir", bg="#F0F0F0", fg="#444444", bd=0,
            font=("Segoe UI", 9), cursor="hand2",
            padx=10, pady=6, command=self._open,
        ).pack(side="left", padx=(0, 6))
        tk.Button(
            path_row, text="Parcourir…", bg=green, fg="white", bd=0,
            font=("Segoe UI", 9, "bold"), cursor="hand2",
            padx=10, pady=6, activebackground=accent,
            activeforeground="white", command=self._browse,
        ).pack(side="left")

        tk.Label(
            body,
            text="Les nouvelles sessions seront enregistrées à cet emplacement.\n"
                 "L'ancien fichier n'est pas déplacé automatiquement.",
            bg=bg, fg=muted, font=("Segoe UI", 8), justify="left",
        ).pack(anchor="w", pady=(0, 10))

        btn_row = tk.Frame(body, bg=bg)
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
            btn_row, text="Enregistrer", bg=accent, fg="white", bd=0,
            font=("Segoe UI", 9, "bold"), cursor="hand2", padx=12, pady=6,
            activebackground=green, activeforeground="white",
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
                messagebox.showerror(
                    "Erreur", f"Impossible de créer le dossier :\n{exc}", parent=self,
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
    ) -> None:
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

    def _build(self) -> None:
        bg = "#FFFFFF"
        accent = "#40916C"
        green = "#52B788"
        red = "#E63946"
        muted = "#888888"
        border = "#E0E0E0"

        hdr = tk.Frame(self, bg=accent, height=36)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="✏  Gérer les projets", bg=accent, fg="white",
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left", padx=12, pady=6)

        body = tk.Frame(self, bg=bg, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        top_row = tk.Frame(body, bg=bg)
        top_row.pack(fill="x", pady=(0, 6))
        tk.Label(
            top_row, text="Projet", bg=bg, fg=muted, font=("Segoe UI", 8, "bold"),
        ).pack(side="left")

        sort_combo = ttk.Combobox(
            top_row, textvariable=self.sort_mode, state="readonly",
            values=[label for label, _mode in SORT_OPTIONS],
            font=("Segoe UI", 8), width=8,
        )
        sort_combo.pack(side="right", padx=(8, 0))
        sort_combo.bind("<<ComboboxSelected>>", self._on_sort_selected)

        self.project_picker = SearchableDropdown(
            body, self.selected_name,
            bg=bg, field_bg="#F5F5F5", border=border,
            muted=muted, accent=green, highlight=green, text="#222222",
            placeholder="Choisir un projet…",
        )
        self.project_picker.pack(fill="x", pady=(0, 12))
        self.bind("<Configure>", self._on_dialog_configure)

        btn_frame = tk.Frame(body, bg=bg)
        btn_frame.pack(fill="x", pady=(0, 0))

        def make_btn(parent: tk.Frame, text: str, color: str, cmd: Callable[[], None]) -> tk.Button:
            button = tk.Button(
                parent, text=text, bg=color, fg="white", bd=0,
                font=("Segoe UI", 9, "bold"), cursor="hand2",
                padx=10, pady=6, activebackground=accent,
                activeforeground="white", command=cmd,
            )
            button.pack(side="left", padx=(0, 6))
            return button

        make_btn(btn_frame, "✏  Renommer", accent, self._rename)
        make_btn(btn_frame, "🗑  Supprimer", red, self._delete)

        tk.Frame(body, bg=border, height=1).pack(fill="x", pady=(12, 8))

        tk.Label(
            body, text="Ajouter un projet", bg=bg, fg=muted, font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", pady=(0, 4))
        add_row = tk.Frame(body, bg=bg)
        add_row.pack(fill="x")
        self.new_var = tk.StringVar()
        entry = tk.Entry(
            add_row, textvariable=self.new_var,
            font=("Segoe UI", 10), bg="#F5F5F5", fg="#222222",
            relief="flat", bd=0, highlightthickness=1,
            highlightbackground=border, highlightcolor=green,
        )
        entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))
        entry.bind("<Return>", lambda _event: self._add())
        tk.Button(
            add_row, text="＋ Ajouter", bg=green, fg="white", bd=0,
            font=("Segoe UI", 9, "bold"), cursor="hand2",
            padx=10, pady=6, activebackground=accent,
            activeforeground="white", command=self._add,
        ).pack(side="left")

        tk.Button(
            body, text="Fermer", bg="#F0F0F0", fg="#444444", bd=0,
            font=("Segoe UI", 9), cursor="hand2", padx=10, pady=6,
            command=self._close,
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
                "Sélection", "Sélectionnez d'abord un projet.", parent=self,
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
            "Renommer", f"Nouveau nom pour « {old_name} » :",
            initialvalue=old_name, parent=self,
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
        self.projects.append({
            "name": name,
            "created": datetime.now().isoformat(timespec="seconds"),
        })
        self.new_var.set("")
        self.selected_name.set(name)
        self._refresh_picker()

    def _close(self) -> None:
        self.destroy()

    def destroy(self) -> None:
        self.project_picker.close_if_open()
        save_config({
            "projects": self.projects,
            "sort_mode": self._current_sort_mode(),
            "excel_file": self.excel_file,
        })
        self.on_close(self.projects, self._current_sort_mode())
        self.grab_release()
        super().destroy()
