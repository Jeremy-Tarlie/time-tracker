import os
import sys

from time_tracker.constants import (
    ACTIVE_SESSION_NAME,
    DEFAULT_EXCEL_NAME,
    INSTANCE_LOCK_NAME,
    PENDING_SESSIONS_NAME,
    PORTABLE_FLAG_NAME,
)

_PACKAGE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(_PACKAGE_DIR)

_portable_forced = False


def is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def get_running_exe() -> str:
    if is_frozen():
        return os.path.abspath(sys.executable)
    return os.path.abspath(sys.argv[0])


def get_exe_dir() -> str:
    return os.path.dirname(get_running_exe())


def is_portable_mode() -> bool:
    if _portable_forced:
        return True
    if "--portable" in sys.argv:
        return True
    flag_path = os.path.join(get_exe_dir(), PORTABLE_FLAG_NAME)
    return os.path.isfile(flag_path)


def set_portable_forced(enabled: bool) -> None:
    global _portable_forced
    _portable_forced = enabled


def get_install_dir() -> str:
    """Dossier d'installation Windows (AppData), hors mode portable."""
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        return os.path.join(base, "TimeTracker")
    if sys.platform == "darwin":
        return os.path.join(
            os.path.expanduser("~"), "Library", "Application Support", "TimeTracker"
        )
    return os.path.join(os.path.expanduser("~"), ".local", "share", "TimeTracker")


def get_install_exe() -> str:
    name = "TimeTracker.exe" if sys.platform == "win32" else "TimeTracker"
    return os.path.join(get_install_dir(), name)


INSTALL_DIR = get_install_dir()
INSTALL_EXE = get_install_exe()


def get_app_data_dir() -> str:
    if is_portable_mode():
        data_dir = get_exe_dir() if is_frozen() else PROJECT_ROOT
        os.makedirs(data_dir, exist_ok=True)
        return data_dir
    if is_frozen():
        os.makedirs(INSTALL_DIR, exist_ok=True)
        return INSTALL_DIR
    return PROJECT_ROOT


def paths_equal(left: str, right: str) -> bool:
    return os.path.normcase(os.path.normpath(left)) == os.path.normcase(os.path.normpath(right))


def refresh_data_paths() -> None:
    """Recalcule les chemins de données (après détection portable)."""
    global APP_DIR, CONFIG_FILE, PENDING_SESSIONS_FILE, ACTIVE_SESSION_FILE, INSTANCE_LOCK_FILE
    APP_DIR = get_app_data_dir()
    CONFIG_FILE = os.path.join(APP_DIR, "projects.json")
    PENDING_SESSIONS_FILE = os.path.join(APP_DIR, PENDING_SESSIONS_NAME)
    ACTIVE_SESSION_FILE = os.path.join(APP_DIR, ACTIVE_SESSION_NAME)
    INSTANCE_LOCK_FILE = os.path.join(APP_DIR, INSTANCE_LOCK_NAME)


APP_DIR = get_app_data_dir()
CONFIG_FILE = os.path.join(APP_DIR, "projects.json")
PENDING_SESSIONS_FILE = os.path.join(APP_DIR, PENDING_SESSIONS_NAME)
ACTIVE_SESSION_FILE = os.path.join(APP_DIR, ACTIVE_SESSION_NAME)
INSTANCE_LOCK_FILE = os.path.join(APP_DIR, INSTANCE_LOCK_NAME)


def default_excel_path() -> str:
    return os.path.join(get_app_data_dir(), DEFAULT_EXCEL_NAME)


def normalize_excel_path(path: str) -> str:
    cleaned = os.path.abspath(os.path.expanduser(path.strip()))
    if not cleaned.lower().endswith(".xlsx"):
        cleaned = f"{cleaned}.xlsx"
    return cleaned
