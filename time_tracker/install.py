import hashlib
import os
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

from time_tracker.constants import APP_NAME, APP_VERSION, REGISTRY_KEY, REGISTRY_NAME
from time_tracker.paths import (
    INSTALL_DIR,
    INSTALL_EXE,
    get_running_exe,
    is_frozen,
    paths_equal,
)


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
    ps_lnk = lnk_path.replace("'", "''")
    ps_target = target.replace("'", "''")
    ps_work = work_dir.replace("'", "''")
    ps_name = APP_NAME.replace("'", "''")
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
