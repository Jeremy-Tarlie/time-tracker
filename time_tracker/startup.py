"""Inscription au démarrage de session : Windows / Linux / macOS."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from time_tracker.constants import APP_NAME, REGISTRY_KEY, REGISTRY_NAME
from time_tracker.paths import is_frozen, is_portable_mode


def startup_available() -> bool:
    if is_portable_mode():
        return False
    return sys.platform in {"win32", "linux", "darwin"}


def resolve_launch_command(exe_path: str | None = None) -> str:
    """Commande à lancer au démarrage (une ligne shell / registre)."""
    if exe_path:
        return f'"{os.path.abspath(exe_path)}"'
    if is_frozen():
        return f'"{os.path.abspath(sys.executable)}"'
    python = os.path.abspath(sys.executable)
    script = os.path.abspath(sys.argv[0])
    return f'"{python}" "{script}"'


def _linux_autostart_path() -> Path:
    base = Path.home() / ".config" / "autostart"
    return base / "timetracker.desktop"


def _macos_plist_path() -> Path:
    return Path.home() / "Library" / "LaunchAgents" / "com.timetracker.app.plist"


def register_startup(exe_path: str | None = None) -> None:
    if not startup_available():
        return
    command = resolve_launch_command(exe_path)
    if sys.platform == "win32":
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            REGISTRY_KEY,
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            winreg.SetValueEx(key, REGISTRY_NAME, 0, winreg.REG_SZ, command)
        return
    if sys.platform == "linux":
        path = _linux_autostart_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            "\n".join(
                [
                    "[Desktop Entry]",
                    "Type=Application",
                    f"Name={APP_NAME}",
                    f"Exec={command}",
                    "X-GNOME-Autostart-enabled=true",
                    "Hidden=false",
                    "",
                ]
            ),
            encoding="utf-8",
        )
        return
    if sys.platform == "darwin":
        path = _macos_plist_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        # Exec = argv list : découpe naive des segments quotés.
        parts = _split_command(command)
        args_xml = "\n".join(f"    <string>{_xml_escape(part)}</string>" for part in parts)
        path.write_text(
            "\n".join(
                [
                    '<?xml version="1.0" encoding="UTF-8"?>',
                    '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" '
                    '"http://www.apple.com/DTDs/PropertyList-1.0.dtd">',
                    '<plist version="1.0">',
                    "<dict>",
                    "  <key>Label</key>",
                    "  <string>com.timetracker.app</string>",
                    "  <key>ProgramArguments</key>",
                    "  <array>",
                    args_xml,
                    "  </array>",
                    "  <key>RunAtLoad</key>",
                    "  <true/>",
                    "</dict>",
                    "</plist>",
                    "",
                ]
            ),
            encoding="utf-8",
        )


def unregister_startup() -> None:
    if is_portable_mode():
        return
    if sys.platform == "win32":
        import winreg

        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                REGISTRY_KEY,
                0,
                winreg.KEY_SET_VALUE,
            ) as key:
                winreg.DeleteValue(key, REGISTRY_NAME)
        except FileNotFoundError:
            return
        except OSError:
            return
        return
    if sys.platform == "linux":
        path = _linux_autostart_path()
        if path.is_file():
            path.unlink()
        return
    if sys.platform == "darwin":
        path = _macos_plist_path()
        if path.is_file():
            path.unlink()


def is_startup_registered() -> bool:
    if is_portable_mode():
        return False
    if sys.platform == "win32":
        import winreg

        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                REGISTRY_KEY,
                0,
                winreg.KEY_READ,
            ) as key:
                winreg.QueryValueEx(key, REGISTRY_NAME)
                return True
        except FileNotFoundError:
            return False
        except OSError:
            return False
    if sys.platform == "linux":
        return _linux_autostart_path().is_file()
    if sys.platform == "darwin":
        return _macos_plist_path().is_file()
    return False


def apply_startup_preference(enabled: bool, exe_path: str | None = None) -> None:
    if enabled:
        register_startup(exe_path)
    else:
        unregister_startup()


def startup_label() -> str:
    if sys.platform == "win32":
        return "Lancer au démarrage de Windows"
    if sys.platform == "darwin":
        return "Lancer au démarrage de macOS"
    if sys.platform == "linux":
        return "Lancer au démarrage de la session"
    return "Lancer au démarrage"


def _xml_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    )


def _split_command(command: str) -> list[str]:
    parts: list[str] = []
    current: list[str] = []
    in_quote = False
    for char in command.strip():
        if char == '"':
            in_quote = not in_quote
            continue
        if char.isspace() and not in_quote:
            if current:
                parts.append("".join(current))
                current = []
            continue
        current.append(char)
    if current:
        parts.append("".join(current))
    return parts
