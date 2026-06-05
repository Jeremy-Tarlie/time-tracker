import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import time
import threading
from datetime import datetime
import os
import json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_FILE  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "heures_projet.xlsx")
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "projects.json")

# ─── CONFIG ────────────────────────────────────────────────────────────────────

def load_projects():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return ["Projet 1", "Projet 2", "Projet 3"]

def save_projects(projects):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)

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
            p = row[1]; h = row[4] or 0
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

    def __init__(self, parent, projects, on_close):
        super().__init__(parent)
        self.projects  = list(projects)
        self.on_close  = on_close
        self.title("Gérer les projets")
        self.resizable(False, False)
        self.grab_set()
        self.attributes("-topmost", True)
        self.configure(bg="#FFFFFF")
        self._build()
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

        # Header
        hdr = tk.Frame(self, bg=ACCENT, height=36)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="✏  Gérer les projets", bg=ACCENT, fg="white",
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=12, pady=6)

        body = tk.Frame(self, bg=BG, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Vos projets", bg=BG, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

        # Listbox + scrollbar
        list_frame = tk.Frame(body, bg=BORDER, bd=1, relief="flat")
        list_frame.pack(fill="both", expand=True)
        sb = tk.Scrollbar(list_frame, orient="vertical")
        self.lb = tk.Listbox(
            list_frame, yscrollcommand=sb.set, selectmode="single",
            font=("Segoe UI", 10), bg=BG, fg="#222222",
            selectbackground=GREEN, selectforeground="white",
            activestyle="none", relief="flat", bd=0,
            highlightthickness=0, width=28, height=10
        )
        sb.config(command=self.lb.yview)
        sb.pack(side="right", fill="y")
        self.lb.pack(fill="both", expand=True, padx=2, pady=2)
        for p in self.projects:
            self.lb.insert("end", p)

        # Action buttons
        btn_frame = tk.Frame(body, bg=BG)
        btn_frame.pack(fill="x", pady=(10, 0))

        def btn(parent, text, color, cmd):
            b = tk.Button(parent, text=text, bg=color, fg="white", bd=0,
                          font=("Segoe UI", 9, "bold"), cursor="hand2",
                          padx=10, pady=6, activebackground=ACCENT,
                          activeforeground="white", command=cmd)
            b.pack(side="left", padx=(0, 6))
            return b

        btn(btn_frame, "✏  Renommer", ACCENT,  self._rename)
        btn(btn_frame, "🗑  Supprimer", RED,    self._delete)

        # Separator
        tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=(12, 8))

        # Add new project inline
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
        entry.bind("<Return>", lambda e: self._add())
        tk.Button(add_row, text="＋ Ajouter", bg=GREEN, fg="white", bd=0,
                  font=("Segoe UI", 9, "bold"), cursor="hand2",
                  padx=10, pady=6, activebackground=ACCENT,
                  activeforeground="white", command=self._add).pack(side="left")

        # Close
        tk.Button(body, text="Fermer", bg="#F0F0F0", fg="#444444", bd=0,
                  font=("Segoe UI", 9), cursor="hand2", padx=10, pady=6,
                  command=self._close).pack(anchor="e", pady=(12, 0))

    def _selected_index(self):
        sel = self.lb.curselection()
        if not sel:
            messagebox.showinfo("Sélection", "Sélectionnez d'abord un projet dans la liste.", parent=self)
            return None
        return sel[0]

    def _rename(self):
        idx = self._selected_index()
        if idx is None:
            return
        old = self.projects[idx]
        new = simpledialog.askstring("Renommer", f"Nouveau nom pour « {old} » :", initialvalue=old, parent=self)
        if not new or not new.strip() or new.strip() == old:
            return
        new = new.strip()
        if new in self.projects:
            messagebox.showwarning("Doublon", f"« {new} » existe déjà.", parent=self)
            return
        self.projects[idx] = new
        self.lb.delete(idx)
        self.lb.insert(idx, new)
        self.lb.selection_set(idx)
        rename_project_in_excel(old, new)

    def _delete(self):
        idx = self._selected_index()
        if idx is None:
            return
        name = self.projects[idx]
        if not messagebox.askyesno("Confirmer", f"Supprimer « {name} » ?\n\nLes sessions Excel seront conservées.", parent=self):
            return
        self.projects.pop(idx)
        self.lb.delete(idx)

    def _add(self):
        name = self.new_var.get().strip()
        if not name:
            return
        if name in self.projects:
            messagebox.showwarning("Doublon", f"« {name} » existe déjà.", parent=self)
            return
        self.projects.append(name)
        self.lb.insert("end", name)
        self.lb.selection_clear(0, "end")
        self.lb.selection_set("end")
        self.lb.see("end")
        self.new_var.set("")

    def _close(self):
        save_projects(self.projects)
        self.on_close(self.projects)
        self.destroy()

    def destroy(self):
        save_projects(self.projects)
        self.on_close(self.projects)
        super().destroy()

# ─── MAIN APP ──────────────────────────────────────────────────────────────────

class TimeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("⏱ Time Tracker")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.98)

        self.projects         = load_projects()
        self.selected_project = tk.StringVar(value=self.projects[0] if self.projects else "")
        self.tracking         = False
        self.start_time       = None
        self.elapsed          = 0
        self._running         = False

        self._build_ui()
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
        top_row.pack(fill="x", pady=(0, 4))
        tk.Label(top_row, text="PROJET", bg=BG, fg=MUTED,
                 font=("Segoe UI", 7, "bold")).pack(side="left")
        tk.Button(top_row, text="Gérer ›", bg=BG, fg=GREEN, bd=0,
                  font=("Segoe UI", 7, "underline"), cursor="hand2",
                  activebackground=BG, activeforeground=ACCENT,
                  command=self._open_manage).pack(side="right")

        # Combobox row
        combo_row = tk.Frame(body, bg=BG)
        combo_row.pack(fill="x", pady=(0, 12))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("W.TCombobox",
                        fieldbackground=SUBTLE, background=SUBTLE,
                        foreground=TEXT, selectbackground=GREEN,
                        selectforeground="white", arrowcolor=ACCENT,
                        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
        style.map("W.TCombobox", fieldbackground=[("readonly", SUBTLE)])

        self.combo = ttk.Combobox(combo_row, textvariable=self.selected_project,
                                   values=self.projects, state="readonly",
                                   font=("Segoe UI", 10), style="W.TCombobox", width=22)
        self.combo.pack(side="left", fill="x", expand=True)

        # Thin separator
        tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=(0, 12))

        # Timer
        self.timer_label = tk.Label(body, text="00:00:00", bg=BG, fg=TIMER_COLOR,
                                     font=("Segoe UI Mono", 32, "bold"))
        self.timer_label.pack()

        self.status_label = tk.Label(body, text="● En attente", bg=BG, fg=MUTED,
                                      font=("Segoe UI", 9))
        self.status_label.pack(pady=(2, 14))

        # Start / Stop
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

        # Thin separator
        tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=(12, 8))

        # Open Excel
        tk.Button(body, text="📊  Ouvrir le fichier Excel", bg=BG, fg=MUTED,
                  font=("Segoe UI", 8), bd=0, cursor="hand2",
                  activebackground=BG, activeforeground=ACCENT,
                  command=self._open_excel).pack()

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
        if not self.selected_project.get():
            messagebox.showwarning("Projet manquant", "Sélectionnez un projet d'abord.")
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
        ManageProjectsDialog(self.root, self.projects, self._on_projects_updated)

    def _on_projects_updated(self, new_projects):
        self.projects = new_projects
        self.combo["values"] = self.projects
        if self.selected_project.get() not in self.projects:
            self.selected_project.set(self.projects[0] if self.projects else "")

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
    ensure_excel()
    root = tk.Tk()
    app  = TimeTrackerApp(root)
    root.mainloop()