import hashlib
import os
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

from time_tracker.constants import APP_NAME, APP_VERSION
from time_tracker.paths import (
    INSTALL_DIR,
    INSTALL_EXE,
    get_running_exe,
    is_frozen,
    is_portable_mode,
    paths_equal,
)
from time_tracker.startup import (
    apply_startup_preference,
    is_startup_registered,
    register_startup,
    startup_label,
    unregister_startup,
)
from time_tracker.theme import (
    THEME,
    apply_window,
    dialog_header,
    font_ui,
    make_button,
    style_checkbutton,
)

__all__ = [
    "apply_startup_preference",
    "ensure_installed",
    "is_startup_registered",
    "register_startup",
    "unregister_startup",
]


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
    for filename in (
        "heures_projet.xlsx",
        "projects.json",
        "pending_sessions.json",
        "active_session.json",
    ):
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


def _ask_install_options(first_install: bool) -> tuple[bool, bool]:
    """Retourne (create_shortcuts, launch_at_startup)."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    dialog = tk.Toplevel(root)
    dialog.title(APP_NAME)
    apply_window(dialog)
    dialog.resizable(False, False)
    dialog.grab_set()
    dialog.attributes("-topmost", True)

    create_var = tk.BooleanVar(value=True)
    startup_var = tk.BooleanVar(value=False)
    result: dict[str, bool] = {"ok": True}
    t = THEME
    title = "Installation" if first_install else f"Mise à jour v{APP_VERSION}"
    dialog_header(dialog, title)
    body = tk.Frame(dialog, bg=t.bg, padx=16, pady=14)
    body.pack(fill="both", expand=True)
    tk.Label(
        body,
        text=f"Dossier : {INSTALL_DIR}",
        bg=t.bg,
        fg=t.muted,
        font=font_ui(9),
        wraplength=360,
        justify="left",
    ).pack(anchor="w", pady=(0, 10))
    create_cb = tk.Checkbutton(
        body,
        text="Créer les raccourcis Bureau et Menu Démarrer",
        variable=create_var,
    )
    style_checkbutton(create_cb)
    create_cb.pack(anchor="w")
    startup_cb = tk.Checkbutton(
        body,
        text=startup_label(),
        variable=startup_var,
    )
    style_checkbutton(startup_cb)
    startup_cb.pack(anchor="w", pady=(4, 12))

    def confirm() -> None:
        result["ok"] = True
        dialog.destroy()

    make_button(body, "Continuer", confirm, variant="primary", font_size=9).pack(anchor="e")

    dialog.update_idletasks()
    dialog.geometry("400x240")
    dialog.wait_window()
    root.destroy()
    if not result["ok"]:
        return True, False
    return bool(create_var.get()), bool(startup_var.get())


def _persist_startup_preference(launch_at_startup: bool) -> None:
    from time_tracker.config import load_config, save_config
    from time_tracker.paths import CONFIG_FILE

    # CONFIG_FILE pointe encore vers l'ancien data dir : écrire dans INSTALL_DIR.
    config_path = os.path.join(INSTALL_DIR, "projects.json")
    if os.path.exists(config_path):
        cfg = load_config(config_path)
    else:
        cfg = load_config(CONFIG_FILE) if os.path.exists(CONFIG_FILE) else load_config()
    cfg["launch_at_startup"] = launch_at_startup
    save_config(cfg, config_path)


def ensure_installed() -> None:
    """Installe ou met à jour sous Windows (opt-in démarrage / raccourcis)."""
    if not is_frozen() or os.name != "nt":
        return
    if is_portable_mode():
        return

    os.makedirs(INSTALL_DIR, exist_ok=True)
    running_exe = get_running_exe()

    if paths_equal(running_exe, INSTALL_EXE):
        return

    first_install = not os.path.exists(INSTALL_EXE)
    create_shortcuts, launch_at_startup = _ask_install_options(first_install)

    if first_install:
        _migrate_data(os.path.dirname(running_exe), INSTALL_DIR)
    needs_update = not first_install and _file_hash(running_exe) != _file_hash(INSTALL_EXE)
    shutil.copy2(running_exe, INSTALL_EXE)
    _persist_startup_preference(launch_at_startup)

    if launch_at_startup:
        register_startup(INSTALL_EXE)
    else:
        unregister_startup()

    if create_shortcuts:
        desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop", f"{APP_NAME}.lnk")
        start_menu = os.path.join(
            os.environ.get("APPDATA", ""),
            "Microsoft",
            "Windows",
            "Start Menu",
            "Programs",
            f"{APP_NAME}.lnk",
        )
        _create_shortcut(desktop, INSTALL_EXE, INSTALL_DIR)
        _create_shortcut(start_menu, INSTALL_EXE, INSTALL_DIR)

    root = tk.Tk()
    root.withdraw()
    if first_install:
        message = (
            f"Installation terminée.\n\n"
            f"Dossier : {INSTALL_DIR}\n"
            + ("Raccourcis créés.\n" if create_shortcuts else "")
            + (
                "Démarrage Windows activé.\n"
                if launch_at_startup
                else "Démarrage Windows non activé (modifiable dans Réglages).\n"
            )
        )
    elif needs_update:
        message = (
            f"Mise à jour terminée (v{APP_VERSION}).\n\n"
            f"Dossier : {INSTALL_DIR}\n"
            f"L'application va redémarrer avec la nouvelle version."
        )
    else:
        message = f"Lancement depuis l'installation.\n\nDossier : {INSTALL_DIR}"
    messagebox.showinfo(APP_NAME, message, parent=root)
    root.destroy()

    creationflags = subprocess.DETACHED_PROCESS if hasattr(subprocess, "DETACHED_PROCESS") else 0
    subprocess.Popen([INSTALL_EXE], close_fds=True, creationflags=creationflags)
    sys.exit(0)
