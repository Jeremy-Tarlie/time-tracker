from __future__ import annotations

import sys
import tkinter as tk
from tkinter import messagebox

from time_tracker.config import load_config_detailed, save_config
from time_tracker.excel import ensure_excel
from time_tracker.install import ensure_installed
from time_tracker.instance import ensure_single_instance
from time_tracker.logging_setup import setup_logging
from time_tracker.paths import refresh_data_paths
from time_tracker.updater import try_startup_auto_update


def _enable_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        awareness = getattr(ctypes.windll.shcore, "SetProcessDpiAwareness", None)
        if awareness is not None:
            awareness(1)
            return
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        return


def main() -> None:
    _enable_dpi_awareness()
    refresh_data_paths()
    setup_logging()

    ensure_installed()
    refresh_data_paths()

    if not ensure_single_instance():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Time Tracker",
            "Time Tracker est déjà en cours d'exécution.",
            parent=root,
        )
        root.destroy()
        return

    try_startup_auto_update()

    result = load_config_detailed()

    if result.was_corrupt:
        warn = tk.Tk()
        warn.withdraw()
        backup = result.corrupt_backup_path or "(non déplacé)"
        messagebox.showwarning(
            "Configuration corrompue",
            "projects.json était illisible et a été mis de côté.\n\n"
            f"Sauvegarde : {backup}\n\n"
            "Une configuration par défaut va être utilisée.",
            parent=warn,
        )
        warn.destroy()

    if result.should_persist:
        save_config(result.config)

    ensure_excel(result.config["excel_file"])

    root = tk.Tk()
    from time_tracker.app import TimeTrackerApp

    TimeTrackerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
