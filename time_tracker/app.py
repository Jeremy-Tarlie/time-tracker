from __future__ import annotations

import os
import re
import sys
import threading
import time
import tkinter as tk
from collections.abc import Callable
from datetime import datetime
from tkinter import messagebox, ttk
from typing import Literal

from time_tracker.config import (
    load_config,
    project_names,
    save_config,
    sort_projects,
    touch_project_last_used,
)
from time_tracker.constants import (
    DEFAULT_SORT,
    SORT_LABELS,
    SORT_OPTIONS,
    UPDATE_CHECK_INTERVAL_MS,
    HotkeyAction,
    HotkeysConfig,
    PendingSession,
    ProjectEntry,
    SortMode,
)
from time_tracker.dialogs import (
    AboutDialog,
    ExcelFileDialog,
    ManageProjectsDialog,
    ManualSessionDialog,
    PeriodHistoryDialog,
    SettingsDialog,
    StartAtDialog,
    TodayHistoryDialog,
    UpdatePromptDialog,
)
from time_tracker.excel import (
    ExcelWriteError,
    ensure_excel,
    format_session_duration,
    open_excel_file,
)
from time_tracker.hotkeys import (
    HOTKEY_ACTIONS,
    GlobalHotkeyController,
    format_hotkey_display,
    hotkey_to_tk,
    parse_hotkeys,
)
from time_tracker.instance import release_instance_lock
from time_tracker.logging_setup import get_logger
from time_tracker.paths import is_frozen
from time_tracker.session_store import (
    clear_active_session,
    enqueue_pending_session,
    load_active_session,
    load_pending_sessions,
)
from time_tracker.session_writer import flush_pending_sync, run_in_thread, write_session_sync
from time_tracker.startup import (
    apply_startup_preference,
    is_startup_registered,
    startup_available,
    startup_label,
)
from time_tracker.theme import (
    THEME,
    Theme,
    apply_ttk_theme,
    apply_window,
    font_mono,
    font_ui,
    hairline,
    make_button,
    section_label,
)
from time_tracker.timer_controller import TimerState, format_elapsed
from time_tracker.tray import TrayController
from time_tracker.updater import (
    UpdateAvailability,
    UpdateCheckResult,
    check_for_update,
    decide_apply_prompt,
    download_update,
    prepare_and_relaunch,
)
from time_tracker.widgets import SearchableDropdown

logger = get_logger()

_PENDING_RETRY_MS = 15_000
_GEOMETRY_DEBOUNCE_MS = 400
_MIN_WIDTH = 300
_MIN_HEIGHT = 220
_TIMER_FONT_SIZE = 32


class TimeTrackerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Time Tracker")
        self.root.minsize(_MIN_WIDTH, _MIN_HEIGHT)
        self.root.resizable(True, True)
        self._theme: Theme = THEME
        apply_window(self.root, self._theme)

        config = load_config()
        self.projects: list[ProjectEntry] = config["projects"]
        self.excel_file: str = config["excel_file"]
        self.launch_at_startup = bool(config.get("launch_at_startup", False))
        self.always_on_top = bool(config.get("always_on_top", True))
        self.window_geometry: str | None = config.get("window_geometry")
        self.ui_compact = bool(config.get("ui_compact", False))
        self.hotkeys: HotkeysConfig = parse_hotkeys(config.get("hotkeys"))
        self.sort_mode = tk.StringVar(value=SORT_LABELS[config["sort_mode"]])
        self._apply_topmost()
        names = project_names(sort_projects(self.projects, config["sort_mode"]))
        self.selected_project = tk.StringVar(value=names[0] if names else "")
        self._timer = TimerState()
        self._tray: TrayController | None = None
        self._quitting = False
        self._geometry_job: str | None = None
        self._pending_retry_job: str | None = None
        self._global_hotkeys: GlobalHotkeyController | None = None
        self._tk_hotkey_bindings: list[str] = []
        self._hotkey_permission_warned = False

        self._body: tk.Frame | None = None
        self._top_row: tk.Frame | None = None
        self._details_frame: tk.Frame | None = None
        self._details_toggle: tk.Button | None = None
        self._manage_btn: tk.Button | None = None
        self._pause_btn: tk.Button | None = None
        self._retry_btn: tk.Button | None = None
        self._manual_link: tk.Button | None = None
        self._start_at_btn: tk.Button | None = None
        self._scroll_canvas: tk.Canvas | None = None
        self._scroll_bar: ttk.Scrollbar | None = None
        self._scroll_window: int | None = None
        self._sessions_divider: tk.Frame | None = None
        self._status_dot: tk.Label | None = None
        self._status_wrap: tk.Frame | None = None
        self._update_btn: tk.Button | None = None
        self._pending_update: UpdateAvailability | None = None
        self._update_check_job: str | None = None
        self._update_apply_job: str | None = None
        self._update_applying = False

        self._build_ui()
        self._setup_hotkeys()
        self._refresh_project_list()
        self._restore_or_position_window()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_request)
        self.root.after(200, self._post_init)

    @property
    def tracking(self) -> bool:
        return self._timer.tracking

    @property
    def paused(self) -> bool:
        return self._timer.paused

    @property
    def active_project(self) -> str | None:
        return self._timer.active_project

    @property
    def start_time(self) -> datetime | None:
        return self._timer.start_time

    @property
    def elapsed(self) -> float:
        return self._timer.elapsed

    @property
    def _writing(self) -> bool:
        return self._timer.writing

    @_writing.setter
    def _writing(self, value: bool) -> None:
        self._timer.writing = value

    @property
    def _running(self) -> bool:
        return self._timer.running

    @property
    def _timer_gen(self) -> int:
        return self._timer.timer_gen

    @property
    def _accumulated(self) -> float:
        return self._timer.accumulated

    def _post_init(self) -> None:
        self._setup_tray()
        self._recover_active_session()
        self._flush_pending_async(silent_if_empty=True, quiet=True)
        self._schedule_pending_retry()
        if is_frozen():
            self._start_update_checks()

    def _setup_tray(self) -> None:
        def schedule_show() -> None:
            self.root.after(0, self._show_window)

        def schedule_quit() -> None:
            self.root.after(0, self._quit_app)

        def schedule_toggle() -> None:
            self.root.after(0, self._tray_toggle_timer)

        self._tray = TrayController(
            on_show=schedule_show,
            on_quit=schedule_quit,
            on_toggle_timer=schedule_toggle,
        )
        self._tray.start()

    def _show_window(self) -> None:
        self.root.deiconify()
        self._apply_topmost()
        self.root.lift()
        self.root.focus_force()

    def _tray_toggle_timer(self) -> None:
        if self.tracking and not self.paused:
            self._stop()
        elif not self.tracking:
            self._start()

    def _default_corner_geometry(self) -> None:
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        self.root.geometry(f"+{sw - w - 30}+{sh - h - 60}")

    def _clamp_geometry(self, geometry: str) -> str | None:
        """Parse WxH±X±Y et clamp dans l'écran ; None si invalide."""
        match = re.fullmatch(
            r"(?P<w>\d+)x(?P<h>\d+)(?P<x>[+-]\d+)(?P<y>[+-]\d+)",
            geometry.strip(),
        )
        if match is None:
            return None
        width = int(match.group("w"))
        height = int(match.group("h"))
        x = int(match.group("x"))
        y = int(match.group("y"))
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        width = max(_MIN_WIDTH, min(width, sw))
        height = max(_MIN_HEIGHT, min(height, sh))
        x = max(0, min(x, max(0, sw - width)))
        y = max(0, min(y, max(0, sh - height)))
        return f"{width}x{height}+{x}+{y}"

    def _restore_or_position_window(self) -> None:
        if self.window_geometry:
            clamped = self._clamp_geometry(self.window_geometry)
            if clamped:
                self.root.geometry(clamped)
                return
        if self.ui_compact:
            self.root.update_idletasks()
            default_h = max(_MIN_HEIGHT, int(self.root.winfo_reqheight()))
        else:
            default_h = 560
        self.root.geometry(f"{_MIN_WIDTH}x{default_h}")
        self._default_corner_geometry()

    def _capture_geometry(self) -> None:
        if self._quitting:
            return
        geom = self.root.geometry()
        if geom == self.window_geometry:
            return
        self.window_geometry = geom
        self._save_config()

    def _schedule_geometry_save(self) -> None:
        if self._geometry_job is not None:
            self.root.after_cancel(self._geometry_job)
        self._geometry_job = self.root.after(_GEOMETRY_DEBOUNCE_MS, self._capture_geometry)

    def _apply_topmost(self) -> None:
        self.root.attributes("-topmost", bool(self.always_on_top))

    def _set_status(
        self,
        text: str,
        *,
        kind: Literal["idle", "running", "paused", "busy", "ok", "error"] = "idle",
    ) -> None:
        t = self._theme
        colors: dict[str, tuple[str, str]] = {
            "idle": (t.muted, t.field),
            "running": (t.success, t.pill_ok_bg),
            "paused": (t.warning_dim, t.pill_paused_bg),
            "busy": (t.muted, t.field),
            "ok": (t.success, t.pill_ok_bg),
            "error": (t.danger, t.pill_error_bg),
        }
        fg, pill_bg = colors[kind]
        if self._status_wrap is not None:
            self._status_wrap.configure(bg=pill_bg)
        if self._status_dot is not None:
            self._status_dot.configure(bg=pill_bg, fg=fg)
        self.status_label.configure(bg=pill_bg, fg=fg, text=text)

    def _hotkey_handlers(self) -> dict[HotkeyAction, Callable[[], None]]:
        return {
            "toggle": self._hotkey_toggle,
            "stop": self._hotkey_stop,
            "settings": self._hotkey_settings,
            "show": self._show_window,
            "today": self._open_today,
            "period": self._open_period,
            "manual": self._open_manual,
        }

    def _setup_hotkeys(self) -> None:
        self._unbind_tk_hotkeys()
        handlers = self._hotkey_handlers()
        if self._global_hotkeys is None:

            def _schedule(callback: Callable[[], None]) -> None:
                self.root.after(0, callback)

            self._global_hotkeys = GlobalHotkeyController(
                schedule=_schedule,
                handlers=handlers,
            )
        error = self._global_hotkeys.start(self.hotkeys)
        if error:
            logger.warning("Hotkeys globaux indisponibles (%s) — repli Tk", error)
            self._bind_tk_hotkeys()
            if not self._hotkey_permission_warned and "Permissions macOS" in error:
                self._hotkey_permission_warned = True
                messagebox.showwarning(
                    "Raccourcis globaux",
                    "Les raccourcis globaux sont désactivés faute de permissions macOS.\n\n"
                    "Réglages système → Confidentialité et sécurité →\n"
                    "Accessibilité et Surveillance des entrées : autorisez Time Tracker "
                    "(ou Terminal / Python en développement), puis redémarrez l'app.\n\n"
                    "En attendant, les raccourcis fonctionnent uniquement "
                    "quand la fenêtre est active.",
                    parent=self.root,
                )

    def _unbind_tk_hotkeys(self) -> None:
        for sequence in self._tk_hotkey_bindings:
            self.root.unbind(sequence)
        self._tk_hotkey_bindings = []

    def _bind_tk_hotkeys(self) -> None:
        handlers = self._hotkey_handlers()
        for action in HOTKEY_ACTIONS:
            sequence = hotkey_to_tk(self.hotkeys[action])
            callback = handlers[action]

            def _make(cb: Callable[[], None] = callback) -> Callable[[object], str]:
                def handler(_event: object = None) -> str:
                    cb()
                    return "break"

                return handler

            self.root.bind(sequence, _make())
            self._tk_hotkey_bindings.append(sequence)

    def _hotkey_toggle(self, _event: object = None) -> None:
        if self._writing:
            return
        if self.tracking:
            self._toggle_pause()
        else:
            self._start()

    def _hotkey_stop(self, _event: object = None) -> None:
        if self.tracking:
            self._stop()

    def _hotkey_settings(self, _event: object = None) -> None:
        self._open_settings()

    def _schedule_pending_retry(self) -> None:
        if self._pending_retry_job is not None:
            self.root.after_cancel(self._pending_retry_job)
        self._pending_retry_job = self.root.after(
            _PENDING_RETRY_MS,
            self._auto_retry_pending,
        )

    def _auto_retry_pending(self) -> None:
        self._pending_retry_job = None
        if not self._quitting and load_pending_sessions():
            if not self._writing and not self.tracking:
                self._flush_pending_async(silent_if_empty=True, quiet=True)
        if not self._quitting:
            self._schedule_pending_retry()

    def _build_ui(self) -> None:
        t = self._theme
        self._BG = t.bg
        self._GREEN = t.accent
        self._RED = t.danger
        self._TEXT = t.text
        self._MUTED = t.muted
        self._ACCENT = t.accent
        self._SUBTLE = t.field
        self._BORDER = t.border
        self._TIMER = t.timer
        self._AMBER = t.warning
        self._ON_ACCENT = t.on_accent
        self._ON_SOLID = t.on_solid
        self._SUCCESS = t.success
        self._SURFACE = t.surface

        apply_window(self.root, t)
        apply_ttk_theme()

        brand = tk.Frame(self.root, bg=t.surface)
        brand.pack(fill="x")
        tk.Frame(brand, bg=t.accent, width=4).pack(side="left", fill="y")
        settings_btn = make_button(
            brand,
            "Réglages",
            self._open_settings,
            variant="link",
            font_size=9,
            bold=False,
            padx=10,
            pady=4,
        )
        settings_btn.configure(fg=t.accent, activeforeground=t.accent_dim)
        settings_btn.pack(side="right", padx=(0, 12), pady=12)
        self._update_btn = make_button(
            brand,
            "Mise à jour",
            self._on_update_clicked,
            variant="link",
            font_size=9,
            bold=False,
            padx=10,
            pady=4,
        )
        self._update_btn.configure(fg=t.warning, activeforeground=t.warning)
        # Caché jusqu'à détection d'une version plus récente (frozen).
        brand_text = tk.Frame(brand, bg=t.surface)
        brand_text.pack(side="left", fill="both", expand=True, padx=(14, 8), pady=12)
        tk.Label(
            brand_text,
            text="Time Tracker",
            bg=t.surface,
            fg=t.text,
            font=font_ui(14, bold=True),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            brand_text,
            text="Suivi du temps par projet",
            bg=t.surface,
            fg=t.muted,
            font=font_ui(8),
            anchor="w",
        ).pack(fill="x", pady=(2, 0))
        hairline(self.root, t)

        body = self._build_scrollable_body(t)
        self._body = body

        self._top_row = tk.Frame(body, bg=t.bg)
        self._top_row.pack(fill="x", pady=(0, 8))
        section_label(self._top_row, "Projet").pack(side="left")

        sort_combo = ttk.Combobox(
            self._top_row,
            textvariable=self.sort_mode,
            state="readonly",
            values=[label for label, _mode in SORT_OPTIONS],
            font=font_ui(8),
            width=8,
            style="App.TCombobox",
        )
        sort_combo.pack(side="right", padx=(8, 0))
        sort_combo.bind("<<ComboboxSelected>>", self._on_sort_changed)

        self._manage_btn = make_button(
            self._top_row,
            "Gérer",
            self._open_manage,
            variant="link",
            font_size=9,
            bold=False,
            padx=6,
            pady=0,
        )
        self._manage_btn.configure(fg=t.accent, activeforeground=t.accent_dim)
        self._manage_btn.pack(side="right")

        self.project_picker = SearchableDropdown(
            body,
            self.selected_project,
            bg=t.bg,
            field_bg=t.surface,
            border=t.border,
            muted=t.muted,
            accent=t.accent_dim,
            highlight=t.accent,
            text=t.text,
        )
        self.project_picker.pack(fill="x", pady=(0, 18))
        self.root.bind("<Configure>", self._on_root_configure)

        timer_card = tk.Frame(
            body,
            bg=t.surface,
            highlightthickness=1,
            highlightbackground=t.border,
        )
        timer_card.pack(fill="x", pady=(0, 16))
        timer_inner = tk.Frame(timer_card, bg=t.surface, padx=16, pady=20)
        timer_inner.pack(fill="x")

        self.timer_label = tk.Label(
            timer_inner,
            text="00:00:00",
            bg=t.surface,
            fg=t.timer,
            font=font_mono(_TIMER_FONT_SIZE, bold=True),
        )
        self.timer_label.pack()

        self._status_wrap = tk.Frame(timer_inner, bg=t.field, padx=10, pady=4)
        self._status_wrap.pack(pady=(10, 0))
        self._status_dot = tk.Label(
            self._status_wrap,
            text="●",
            bg=t.field,
            fg=t.muted,
            font=font_ui(8),
        )
        self._status_dot.pack(side="left", padx=(0, 6))
        self.status_label = tk.Label(
            self._status_wrap,
            text="En attente",
            bg=t.field,
            fg=t.muted,
            font=font_ui(9),
        )
        self.status_label.pack(side="left")

        btn_row = tk.Frame(body, bg=t.bg)
        btn_row.pack(fill="x")

        self.start_btn = make_button(
            btn_row,
            "Démarrer",
            self._start,
            variant="primary",
            padx=12,
            pady=12,
        )
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self._pause_btn = make_button(
            btn_row,
            "Pause",
            self._toggle_pause,
            variant="subtle",
            padx=12,
            pady=12,
        )
        self._pause_btn.configure(state="disabled")
        self._pause_btn.pack(side="left", padx=(0, 6))

        self.stop_btn = make_button(
            btn_row,
            "Arrêter",
            self._stop,
            variant="subtle",
            padx=12,
            pady=12,
        )
        self.stop_btn.configure(state="disabled")
        self.stop_btn.pack(side="left", fill="x", expand=True)

        self._start_at_btn = make_button(
            body,
            "Démarrer à…",
            self._start_at,
            variant="link",
            font_size=9,
            bold=False,
            padx=0,
            pady=3,
        )
        self._start_at_btn.configure(fg=t.accent, activeforeground=t.accent_dim)
        self._start_at_btn.pack(fill="x", anchor="w", pady=(8, 0))

        self._manual_link = make_button(
            body,
            self._manual_link_text(),
            self._open_manual,
            variant="link",
            font_size=9,
            bold=False,
            padx=0,
            pady=3,
        )
        self._manual_link.configure(fg=t.accent, activeforeground=t.accent_dim)
        self._manual_link.pack(fill="x", anchor="w", pady=(4, 0))

        self._retry_btn = make_button(
            body,
            "Réessayer l'enregistrement",
            self._retry_pending,
            variant="subtle",
            font_size=9,
            bold=False,
            padx=10,
            pady=9,
        )
        self._retry_btn.configure(fg=t.danger)
        # Affiché uniquement s'il y a des sessions en file (voir _update_retry_button).

        self._details_toggle = make_button(
            body,
            "▾ Détails",
            self._toggle_ui_compact,
            variant="link",
            font_size=9,
            bold=False,
            padx=0,
            pady=3,
        )
        self._details_toggle.pack(fill="x", anchor="w", pady=(14, 0))

        self._details_frame = tk.Frame(body, bg=t.bg)
        self._details_frame.pack(fill="x")

        self._sessions_divider = tk.Frame(self._details_frame, bg=t.border, height=1)
        self._sessions_divider.pack(fill="x", pady=(4, 14))

        section_label(self._details_frame, "Sessions").pack(anchor="w", pady=(0, 6))
        primary_actions = tk.Frame(self._details_frame, bg=t.bg)
        primary_actions.pack(fill="x", pady=(0, 10))

        def action_link(parent: tk.Frame, label: str, command: Callable[[], None]) -> None:
            make_button(
                parent,
                label,
                command,
                variant="link",
                font_size=9,
                bold=False,
                padx=0,
                pady=3,
            ).pack(fill="x", anchor="w")

        action_link(primary_actions, "Aujourd'hui", self._open_today)
        action_link(primary_actions, "Période…", self._open_period)

        section_label(self._details_frame, "Excel").pack(anchor="w", pady=(4, 6))
        secondary = tk.Frame(self._details_frame, bg=t.bg)
        secondary.pack(fill="x")
        action_link(secondary, "Ouvrir Excel", self._open_excel)
        action_link(secondary, "Changer l'emplacement…", self._open_excel_settings)
        action_link(secondary, "À propos…", self._open_about)

        self._set_ui_compact(self.ui_compact, persist=False, resize=False)
        self._update_retry_button()
        self._sync_start_enabled()
        self._set_status("En attente", kind="idle")
        self._refresh_scroll_region()

    def _build_scrollable_body(self, t: Theme) -> tk.Frame:
        """Corps principal dans un Canvas scrollable (fenêtre réduite)."""
        shell = tk.Frame(self.root, bg=t.bg)
        shell.pack(fill="both", expand=True)
        self._scroll_canvas = tk.Canvas(
            shell,
            bg=t.bg,
            highlightthickness=0,
            bd=0,
        )
        self._scroll_bar = ttk.Scrollbar(
            shell,
            orient="vertical",
            command=self._scroll_canvas.yview,
        )
        self._scroll_canvas.configure(yscrollcommand=self._scroll_bar.set)
        self._scroll_bar.pack(side="right", fill="y")
        self._scroll_canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(self._scroll_canvas, bg=t.bg, padx=20, pady=18)
        self._scroll_window = self._scroll_canvas.create_window(
            (0, 0),
            window=body,
            anchor="nw",
        )
        body.bind("<Configure>", self._on_scroll_body_configure)
        self._scroll_canvas.bind("<Configure>", self._on_scroll_canvas_configure)
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        return body

    def _on_scroll_body_configure(self, _event: object = None) -> None:
        self._refresh_scroll_region()

    def _on_scroll_canvas_configure(self, event: tk.Event[tk.Misc]) -> None:
        if self._scroll_canvas is None or self._scroll_window is None:
            return
        self._scroll_canvas.itemconfigure(self._scroll_window, width=int(event.width))

    def _refresh_scroll_region(self) -> None:
        if self._scroll_canvas is None:
            return
        self._scroll_canvas.update_idletasks()
        bbox = self._scroll_canvas.bbox("all")
        if bbox is not None:
            self._scroll_canvas.configure(scrollregion=bbox)

    def _on_mousewheel(self, event: tk.Event[tk.Misc]) -> str | None:
        if self._scroll_canvas is None:
            return None
        try:
            if event.widget.winfo_toplevel() is not self.root:
                return None
        except tk.TclError:
            return None
        self._scroll_canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"

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
        self._sync_start_enabled()

    def _save_config(self) -> None:
        save_config(
            {
                "projects": self.projects,
                "sort_mode": self._current_sort_mode(),
                "excel_file": self.excel_file,
                "launch_at_startup": self.launch_at_startup,
                "always_on_top": self.always_on_top,
                "window_geometry": self.window_geometry,
                "ui_compact": self.ui_compact,
                "hotkeys": self.hotkeys,
            }
        )

    def _toggle_ui_compact(self) -> None:
        self._set_ui_compact(not self.ui_compact, persist=True, resize=True)

    def _set_ui_compact(
        self,
        compact: bool,
        *,
        persist: bool = True,
        resize: bool = True,
    ) -> None:
        self.ui_compact = compact
        if self._details_toggle is not None:
            self._details_toggle.config(text="▸ Détails" if compact else "▾ Détails")

        if compact:
            if self._top_row is not None:
                self._top_row.pack_forget()
            if self._details_frame is not None:
                self._details_frame.pack_forget()
        else:
            if self._top_row is not None and self.project_picker is not None:
                if not bool(self._top_row.winfo_ismapped()):
                    self._top_row.pack(fill="x", pady=(0, 8), before=self.project_picker)
            if self._details_frame is not None and self._details_toggle is not None:
                if not bool(self._details_frame.winfo_ismapped()):
                    self._details_frame.pack(fill="x", after=self._details_toggle)
        self.root.minsize(_MIN_WIDTH, _MIN_HEIGHT)

        self._update_retry_button()
        self._refresh_scroll_region()
        if resize:
            # Hauteur confortable au bascule, puis l'utilisateur peut réduire (scroll).
            floor = 280 if compact else 560
            self.root.update_idletasks()
            needed = max(floor, int(self.root.winfo_reqheight()))
            self._fit_window_height(needed)
        if persist:
            self._save_config()

    def _fit_window_height(self, target_height: int) -> None:
        """Ajuste la hauteur en conservant largeur et position ; sync window_geometry."""
        self.root.update_idletasks()
        width = max(_MIN_WIDTH, self.root.winfo_width())
        current_h = self.root.winfo_height()
        if abs(current_h - target_height) < 8:
            return
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        geom = f"{width}x{target_height}+{x}+{y}"
        self.root.geometry(geom)
        self.window_geometry = geom
        self._schedule_geometry_save()

    def _on_sort_changed(self, _event: object = None) -> None:
        self._refresh_project_list(persist_sort=True)

    def _on_root_configure(self, event: tk.Event[tk.Misc]) -> None:
        if event.widget is self.root:
            self.project_picker.close_if_open()
            self._schedule_geometry_save()

    def _sync_start_enabled(self) -> None:
        if self.tracking or self._writing:
            return
        has_project = bool(project_names(self.projects))
        state: Literal["normal", "disabled"] = "normal" if has_project else "disabled"
        bg = self._GREEN if has_project else self._SUBTLE
        fg = self._ON_ACCENT if has_project else self._MUTED
        self.start_btn.config(state=state, bg=bg, fg=fg)
        if self._start_at_btn is not None:
            self._start_at_btn.config(state=state)

    def _set_project_controls_enabled(self, enabled: bool) -> None:
        self.project_picker.set_enabled(enabled)
        if self._manage_btn is not None:
            self._manage_btn.config(state="normal" if enabled else "disabled")

    def _set_idle_controls(self) -> None:
        self._timer.reset_idle()
        self._writing = False
        self._set_project_controls_enabled(True)
        self.stop_btn.config(state="disabled", bg=self._SUBTLE, fg=self._MUTED)
        if self._pause_btn is not None:
            self._pause_btn.config(
                state="disabled",
                bg=self._SUBTLE,
                fg=self._MUTED,
                text="Pause",
            )
        self._sync_start_enabled()

    def _set_busy_controls(self) -> None:
        self._writing = True
        self.start_btn.config(state="disabled", bg=self._SUBTLE, fg=self._MUTED)
        self.stop_btn.config(state="disabled", bg=self._SUBTLE, fg=self._MUTED)
        if self._pause_btn is not None:
            self._pause_btn.config(state="disabled", bg=self._SUBTLE, fg=self._MUTED)
        if self._start_at_btn is not None:
            self._start_at_btn.config(state="disabled")
        self._set_project_controls_enabled(False)

    def _set_tracking_controls(self) -> None:
        self._writing = False
        self.start_btn.config(state="disabled", bg=self._SUBTLE, fg=self._MUTED)
        self.stop_btn.config(state="normal", bg=self._RED, fg=self._ON_SOLID)
        if self._pause_btn is not None:
            if self.paused:
                self._pause_btn.config(
                    state="normal",
                    bg=self._AMBER,
                    fg=self._ON_SOLID,
                    text="Reprendre",
                )
            else:
                self._pause_btn.config(
                    state="normal",
                    bg=self._SUBTLE,
                    fg=self._TEXT,
                    text="Pause",
                )
        if self._start_at_btn is not None:
            self._start_at_btn.config(state="disabled")
        self._set_project_controls_enabled(False)

    def _format_elapsed(self, seconds: float) -> str:
        return format_elapsed(seconds)

    def _persist_active(self) -> None:
        self._timer.persist()

    def _ask_session_notes(self) -> str:
        dialog = tk.Toplevel(self.root)
        dialog.title("Notes (optionnel)")
        apply_window(dialog, self._theme)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        result: dict[str, str] = {"value": ""}
        body = tk.Frame(dialog, bg=self._BG, padx=16, pady=14)
        body.pack(fill="both", expand=True)
        tk.Label(
            body,
            text="Ajouter une note pour cette session :",
            bg=self._BG,
            fg=self._MUTED,
            font=font_ui(9),
        ).pack(anchor="w", pady=(0, 8))
        note_var = tk.StringVar()
        entry = tk.Entry(
            body,
            textvariable=note_var,
            font=font_ui(10),
            bg=self._SUBTLE,
            fg=self._TEXT,
            insertbackground=self._TEXT,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=self._BORDER,
            highlightcolor=self._GREEN,
        )
        entry.pack(fill="x", ipady=6)
        entry.focus_set()

        def confirm() -> None:
            result["value"] = note_var.get().strip()
            dialog.destroy()

        def skip() -> None:
            result["value"] = ""
            dialog.destroy()

        btn_row = tk.Frame(body, bg=self._BG)
        btn_row.pack(fill="x", pady=(12, 0))
        make_button(btn_row, "Sans note", skip, variant="subtle", font_size=9, bold=False).pack(
            side="left",
        )
        make_button(btn_row, "OK", confirm, variant="primary", font_size=9).pack(side="right")
        entry.bind("<Return>", lambda _event: confirm())
        dialog.bind("<Escape>", lambda _event: skip())

        dialog.update_idletasks()
        px = self.root.winfo_rootx() + max((self.root.winfo_width() - 320) // 2, 0)
        py = self.root.winfo_rooty() + 40
        dialog.geometry(f"320x150+{px}+{py}")
        dialog.wait_window()
        return result["value"]

    def _manual_link_text(self) -> str:
        combo = format_hotkey_display(self.hotkeys["manual"])
        return f"Saisie manuelle… ({combo})"

    def _start(self) -> None:
        if self.tracking or self._writing:
            return
        project = self.selected_project.get().strip()
        if project not in project_names(self.projects):
            messagebox.showwarning(
                "Projet manquant",
                "Sélectionnez un projet valide dans la liste.",
                parent=self.root,
            )
            return
        self._begin_tracking(project, start_at=None)

    def _start_at(self) -> None:
        if self.tracking or self._writing:
            return
        project = self.selected_project.get().strip()
        if project not in project_names(self.projects):
            messagebox.showwarning(
                "Projet manquant",
                "Sélectionnez un projet valide dans la liste.",
                parent=self.root,
            )
            return
        StartAtDialog(
            self.root,
            lambda start_at: self._begin_tracking(project, start_at=start_at),
        )

    def _begin_tracking(self, project: str, *, start_at: datetime | None) -> None:
        if self.tracking or self._writing:
            return
        tick_gen = self._timer.begin(project, now=start_at)
        touch_project_last_used(self.projects, project)
        self._save_config()
        self._refresh_project_list()
        self._set_tracking_controls()
        self._set_status(f"En cours — {project}", kind="running")
        self._persist_active()
        self.timer_label.config(text=self._format_elapsed(self._timer.tick_elapsed()))
        threading.Thread(target=self._tick, args=(tick_gen,), daemon=True).start()
        logger.info("Session démarrée")

    def _toggle_pause(self) -> None:
        if not self.tracking or self.active_project is None:
            return
        if self.paused:
            tick_gen = self._timer.resume_from_pause()
            self._set_tracking_controls()
            self._set_status(f"En cours — {self.active_project}", kind="running")
            self._persist_active()
            threading.Thread(target=self._tick, args=(tick_gen,), daemon=True).start()
            return
        self._timer.pause()
        self._set_tracking_controls()
        self._set_status(f"Pause — {self.active_project}", kind="paused")
        self._persist_active()

    def _apply_timer_text(self, tick_gen: int, text: str) -> None:
        if tick_gen != self._timer_gen:
            return
        self.timer_label.config(text=text)

    def _tick(self, tick_gen: int) -> None:
        while self._running and tick_gen == self._timer_gen and self.start_time is not None:
            elapsed = self._timer.tick_elapsed()
            text = self._format_elapsed(elapsed)
            self.root.after(0, self._apply_timer_text, tick_gen, text)
            if int(elapsed) % 15 == 0:
                self.root.after(0, self._persist_active)
            time.sleep(1)

    def _stop(self) -> None:
        if not self.tracking or self.active_project is None:
            return
        project, session_start, end_time, elapsed = self._timer.stop_clock()

        self._set_busy_controls()
        self._set_status("Enregistrement…", kind="busy")

        notes = self._ask_session_notes()
        self._write_session_async(project, session_start, end_time, notes, elapsed)

    def _write_session_async(
        self,
        project: str,
        session_start: datetime,
        end_time: datetime,
        notes: str,
        elapsed: float,
        *,
        from_pending: bool = False,
    ) -> None:
        def worker() -> None:
            try:
                write_session_sync(project, session_start, end_time, notes, self.excel_file)
            except ExcelWriteError as exc:
                self.root.after(
                    0,
                    self._on_write_failure,
                    project,
                    session_start,
                    end_time,
                    notes,
                    elapsed,
                    str(exc),
                    from_pending,
                )
                return
            self.root.after(
                0,
                self._on_write_success,
                project,
                elapsed,
                from_pending,
            )

        run_in_thread(worker)

    def _on_write_success(
        self,
        project: str,
        elapsed: float,
        from_pending: bool,
    ) -> None:
        self._writing = False
        self._set_idle_controls()
        self.timer_label.config(text="00:00:00")
        duration_label = format_session_duration(elapsed)
        self._set_status("Session enregistrée", kind="ok")
        self._update_retry_button()
        if not from_pending:
            messagebox.showinfo(
                "Session enregistrée",
                f"✔ {project}\nDurée : {duration_label}\n\nEnregistré dans :\n{self.excel_file}\n\n"
                "Si Calc/Excel était ouvert : fermez puis rouvrez le fichier pour voir la ligne.",
                parent=self.root,
            )
        self._set_status("En attente", kind="idle")
        logger.info("Session enregistrée")

    def _on_write_failure(
        self,
        project: str,
        session_start: datetime,
        end_time: datetime,
        notes: str,
        elapsed: float,
        error: str,
        from_pending: bool,
    ) -> None:
        self._writing = False
        if not from_pending:
            enqueue_pending_session(project, session_start, end_time, notes)
        self._set_idle_controls()
        self.timer_label.config(text="00:00:00")
        self._set_status("Non enregistré — en file", kind="error")
        self._update_retry_button()
        messagebox.showerror(
            "Enregistrement impossible",
            f"La session n'a pas été écrite dans Excel.\n"
            f"Elle est conservée localement pour réessai.\n\n{error}",
            parent=self.root,
        )
        if "vrai .xlsx" in error or "illisible ou corrompu" in error:
            self._offer_recreate_excel()
        self._set_status("En attente de réessai", kind="error")
        logger.warning("Échec écriture Excel")

    def _update_retry_button(self) -> None:
        pending = load_pending_sessions()
        if self._retry_btn is None:
            return
        if pending:
            self._retry_btn.config(
                state="normal",
                text=f"↺  Réessayer ({len(pending)})",
            )
            if not bool(self._retry_btn.winfo_ismapped()):
                if self._details_toggle is not None:
                    self._retry_btn.pack(
                        fill="x",
                        pady=(14, 0),
                        before=self._details_toggle,
                    )
                else:
                    self._retry_btn.pack(fill="x", pady=(14, 0))
        else:
            self._retry_btn.pack_forget()

    def _retry_pending(self) -> None:
        self._flush_pending_async(silent_if_empty=False, quiet=False)

    def _flush_pending_async(self, *, silent_if_empty: bool, quiet: bool = False) -> None:
        pending = load_pending_sessions()
        if not pending:
            if not silent_if_empty and not quiet:
                messagebox.showinfo(
                    "File d'attente",
                    "Aucune session en attente.",
                    parent=self.root,
                )
            self._update_retry_button()
            return
        if self._writing or self.tracking:
            return
        self._set_busy_controls()
        self._set_status("Réessai enregistrement…", kind="busy")
        excel_file = self.excel_file

        def worker() -> None:
            remaining, errors = flush_pending_sync(pending, excel_file)
            self.root.after(0, lambda: self._on_flush_done(remaining, errors, quiet=quiet))

        run_in_thread(worker)

    def _on_flush_done(
        self,
        remaining: list[PendingSession],
        errors: list[str],
        *,
        quiet: bool = False,
    ) -> None:
        self._writing = False
        self._set_idle_controls()
        self._update_retry_button()
        if remaining:
            self._set_status("En attente de réessai", kind="error")
            if not quiet:
                detail = errors[0] if errors else "Échec inconnu"
                messagebox.showerror(
                    "Réessai incomplet",
                    f"{len(remaining)} session(s) restent en file.\n\n{detail}",
                    parent=self.root,
                )
            return
        self._set_status("File enregistrée", kind="ok")
        if not quiet:
            messagebox.showinfo(
                "Enregistrement",
                "Toutes les sessions en attente ont été écrites dans Excel.",
                parent=self.root,
            )
        self._set_status("En attente", kind="idle")

    def _recover_active_session(self) -> None:
        active = load_active_session()
        if active is None:
            return
        project = active["project"]
        if project not in project_names(self.projects):
            clear_active_session()
            return
        choice = messagebox.askyesnocancel(
            "Session interrompue",
            f"Une session « {project} » était en cours.\n\n"
            "Oui = reprendre\nNon = clôturer maintenant\nAnnuler = abandonner",
            parent=self.root,
        )
        if choice is None:
            clear_active_session()
            return
        start = datetime.fromisoformat(active["start"])
        accumulated = float(active["accumulated_seconds"])
        if active["paused"]:
            elapsed = accumulated
        else:
            resume_raw = active.get("resume_at")
            if resume_raw:
                try:
                    resume_at = datetime.fromisoformat(resume_raw)
                    elapsed = accumulated + max(
                        0.0,
                        (datetime.now() - resume_at).total_seconds(),
                    )
                except ValueError:
                    elapsed = accumulated + max(
                        0.0,
                        (datetime.now() - start).total_seconds(),
                    )
            else:
                elapsed = accumulated + max(
                    0.0,
                    (datetime.now() - start).total_seconds(),
                )
        if choice is False:
            end_time = datetime.now()
            session_start = datetime.fromtimestamp(end_time.timestamp() - elapsed)
            clear_active_session()
            notes = self._ask_session_notes()
            self._set_busy_controls()
            self._write_session_async(project, session_start, end_time, notes, elapsed)
            return
        self.selected_project.set(project)
        tick_gen = self._timer.restore_running(project, elapsed)
        self.timer_label.config(text=self._format_elapsed(elapsed))
        self._set_tracking_controls()
        self._set_status(f"Reprise — {project}", kind="running")
        self._persist_active()
        threading.Thread(target=self._tick, args=(tick_gen,), daemon=True).start()

    def _start_update_checks(self) -> None:
        self._run_update_check_async()
        self._schedule_update_check()

    def _schedule_update_check(self) -> None:
        if self._quitting:
            return
        if self._update_check_job is not None:
            try:
                self.root.after_cancel(self._update_check_job)
            except tk.TclError as exc:
                logger.warning("update.cancel_job_failed reason=%s", type(exc).__name__)
        self._update_check_job = self.root.after(
            UPDATE_CHECK_INTERVAL_MS,
            self._on_scheduled_update_check,
        )

    def _on_scheduled_update_check(self) -> None:
        self._update_check_job = None
        if self._quitting:
            return
        self._run_update_check_async()
        self._schedule_update_check()

    def _run_update_check_async(self) -> None:
        def worker() -> None:
            result = check_for_update()
            if self._quitting:
                return

            def deliver() -> None:
                self._on_update_check_result(result)

            self.root.after(0, deliver)

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_check_result(self, result: UpdateCheckResult) -> None:
        if self._quitting:
            return
        if result.status == "error":
            # Conserve le bouton / état déjà connu si le réseau échoue.
            return
        if result.status == "available" and result.availability is not None:
            self._pending_update = result.availability
            self._show_update_button(result.availability.remote_version)
            return
        self._pending_update = None
        self._hide_update_button()

    def _show_update_button(self, remote_version: str) -> None:
        btn = self._update_btn
        if btn is None:
            return
        btn.configure(text=f"Mise à jour v{remote_version}")
        if not btn.winfo_ismapped():
            btn.pack(side="right", padx=(0, 4), pady=12)

    def _hide_update_button(self) -> None:
        btn = self._update_btn
        if btn is None:
            return
        if btn.winfo_ismapped():
            btn.pack_forget()

    def _on_update_clicked(self) -> None:
        if self._update_applying:
            return
        availability = self._pending_update
        if availability is None:
            return
        decision = decide_apply_prompt(tracking=self.tracking, writing=self._writing)
        if decision == "ask_user":
            dialog = UpdatePromptDialog(
                self.root,
                remote_version=availability.remote_version,
            )
            if dialog.result != "now":
                return
        self._begin_apply_update()

    def _begin_apply_update(self) -> None:
        if self._quitting or self._update_applying:
            return
        if self._writing:
            self._update_apply_job = self.root.after(200, self._begin_apply_update)
            return
        if self.tracking:
            self._persist_active()
        availability = self._pending_update
        if availability is None:
            return

        self._update_applying = True
        if self._update_btn is not None:
            self._update_btn.configure(state="disabled")

        def worker() -> None:
            try:
                dest = download_update(availability)
            except Exception as exc:
                logger.warning("update.apply_failed reason=%s", type(exc).__name__)
                self.root.after(0, self._on_update_apply_failed)
                return
            if dest is None:
                self.root.after(0, self._on_update_apply_failed)
                return

            def finish() -> None:
                self._finish_apply_update(dest)

            self.root.after(0, finish)

        self._set_status("Mise à jour en cours…", kind="busy")
        threading.Thread(target=worker, daemon=True).start()

    def _finish_apply_update(self, new_path: str) -> None:
        if self._quitting:
            return
        if self._writing:

            def retry() -> None:
                self._finish_apply_update(new_path)

            self._update_apply_job = self.root.after(200, retry)
            return
        if self.tracking:
            self._persist_active()
        try:
            prepare_and_relaunch(new_path)
        except SystemExit:
            raise
        except Exception as exc:
            logger.warning("update.apply_failed reason=%s", type(exc).__name__)
            self._on_update_apply_failed()

    def _on_update_apply_failed(self) -> None:
        self._update_applying = False
        if self._update_btn is not None:
            self._update_btn.configure(state="normal")
        if self._quitting:
            return
        messagebox.showwarning(
            "Mise à jour",
            "Impossible de télécharger ou d'appliquer la mise à jour.\nRéessayez plus tard.",
            parent=self.root,
        )
        self._set_status("Mise à jour échouée", kind="error")

    def _cancel_update_jobs(self) -> None:
        for job_attr in ("_update_check_job", "_update_apply_job"):
            job = getattr(self, job_attr)
            if job is None:
                continue
            try:
                self.root.after_cancel(job)
            except tk.TclError as exc:
                logger.warning("update.cancel_job_failed reason=%s", type(exc).__name__)
            setattr(self, job_attr, None)

    def _on_close_request(self) -> None:
        # Croix = quitter (plus de withdraw silencieux qui laisse des process vivants).
        self._quit_app()

    def _quit_app(self) -> None:
        if self._quitting:
            return
        if self.tracking:
            if not messagebox.askyesno(
                "Quitter",
                "Session en cours : l'arrêter et l'enregistrer avant de quitter ?",
                parent=self.root,
            ):
                self._show_window()
                return
            self._stop()
            # Leave quit to user after write; mark quitting after write is complex —
            # if still tracking after stop kickoff, wait.
            if self._writing:
                self.root.after(500, self._quit_app)
                return
        self._quitting = True
        self._cancel_update_jobs()
        self._capture_geometry()
        if self._global_hotkeys is not None:
            self._global_hotkeys.stop()
        if self._tray is not None:
            self._tray.stop()
        clear_active_session()
        release_instance_lock()
        self.root.destroy()

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
            messagebox.showerror("Création impossible", str(exc), parent=self.root)
            return
        messagebox.showinfo(
            "Fichier prêt",
            f"Nouveau fichier créé :\n{excel_path}\n\n"
            "Utilisez « Réessayer l'enregistrement » pour écrire la session en file.",
            parent=self.root,
        )

    def _open_manage(self) -> None:
        if self.tracking:
            return
        ManageProjectsDialog(
            self.root,
            self.projects,
            self._current_sort_mode(),
            self.excel_file,
            self._on_projects_updated,
            launch_at_startup=self.launch_at_startup,
            always_on_top=self.always_on_top,
            window_geometry=self.window_geometry,
            ui_compact=self.ui_compact,
            hotkeys=self.hotkeys,
        )

    def _on_projects_updated(
        self,
        new_projects: list[ProjectEntry],
        sort_mode: SortMode,
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

    def _open_today(self) -> None:
        TodayHistoryDialog(self.root, self.excel_file, self.projects)

    def _open_period(self) -> None:
        PeriodHistoryDialog(self.root, self.excel_file, self.projects)

    def _open_manual(self) -> None:
        if not self.projects:
            messagebox.showwarning(
                "Projets",
                "Ajoutez un projet avant la saisie manuelle.",
                parent=self.root,
            )
            return
        ManualSessionDialog(self.root, self.projects, self._on_manual_submit)

    def _on_manual_submit(
        self,
        project: str,
        start: datetime,
        end: datetime,
        notes: str,
    ) -> None:
        elapsed = (end - start).total_seconds()
        self._set_busy_controls()
        self._set_status("Enregistrement…", kind="busy")
        self._write_session_async(project, start, end, notes, elapsed)

    def _open_about(self) -> None:
        AboutDialog(self.root, hotkeys=self.hotkeys)

    def _open_settings(self) -> None:
        # Pause l'écoute globale pour ne pas déclencher d'actions pendant la capture.
        if self._global_hotkeys is not None:
            self._global_hotkeys.stop()
        self._unbind_tk_hotkeys()
        dialog = SettingsDialog(
            self.root,
            launch_at_startup=self.launch_at_startup,
            always_on_top=self.always_on_top,
            hotkeys=self.hotkeys,
            startup_available=startup_available(),
            startup_label=startup_label(),
            on_save=self._on_settings_saved,
        )

        def _resume_hotkeys(event: tk.Event[tk.Misc]) -> None:
            if event.widget is dialog and not self._quitting:
                self._setup_hotkeys()

        dialog.bind("<Destroy>", _resume_hotkeys)

    def _on_settings_saved(
        self,
        launch_at_startup: bool,
        always_on_top: bool,
        hotkeys: HotkeysConfig,
    ) -> None:
        self.launch_at_startup = launch_at_startup
        self.always_on_top = always_on_top
        self.hotkeys = parse_hotkeys(hotkeys)
        if self._manual_link is not None:
            self._manual_link.configure(text=self._manual_link_text())
        self._apply_topmost()
        self._save_config()
        if startup_available():
            exe_path = os.path.abspath(sys.executable) if getattr(sys, "frozen", False) else None
            try:
                apply_startup_preference(launch_at_startup, exe_path)
            except OSError as exc:
                messagebox.showwarning(
                    "Démarrage",
                    f"Impossible d'appliquer le démarrage automatique :\n{exc}",
                    parent=self.root,
                )
                logger.warning("Échec apply_startup_preference: %s", exc)
                return
            registered = is_startup_registered()
            if launch_at_startup and not registered:
                messagebox.showwarning(
                    "Démarrage",
                    "La préférence a été enregistrée, mais l'inscription OS a échoué.",
                    parent=self.root,
                )
            logger.info(
                "Préférence démarrage=%s (actif=%s)",
                launch_at_startup,
                registered,
            )
