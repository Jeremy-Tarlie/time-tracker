import tkinter as tk

from time_tracker.app import TimeTrackerApp
from time_tracker.config import load_config, save_config
from time_tracker.excel import ensure_excel
from time_tracker.install import ensure_installed


def main() -> None:
    ensure_installed()
    startup_config = load_config()
    ensure_excel(startup_config["excel_file"])
    save_config(startup_config)
    root = tk.Tk()
    TimeTrackerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
