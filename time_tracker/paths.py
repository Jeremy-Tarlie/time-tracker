import os
import sys

from time_tracker.constants import DEFAULT_EXCEL_NAME

_PACKAGE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(_PACKAGE_DIR)


def is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def get_install_dir() -> str:
    """Dossier de données utilisateur (et d'installation Windows)."""
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
    if is_frozen():
        os.makedirs(INSTALL_DIR, exist_ok=True)
        return INSTALL_DIR
    return PROJECT_ROOT


def get_running_exe() -> str:
    if is_frozen():
        return os.path.abspath(sys.executable)
    return os.path.abspath(sys.argv[0])


def paths_equal(left: str, right: str) -> bool:
    return os.path.normcase(os.path.normpath(left)) == os.path.normcase(os.path.normpath(right))


APP_DIR = get_app_data_dir()
CONFIG_FILE = os.path.join(APP_DIR, "projects.json")


def default_excel_path() -> str:
    return os.path.join(APP_DIR, DEFAULT_EXCEL_NAME)


def normalize_excel_path(path: str) -> str:
    cleaned = os.path.abspath(os.path.expanduser(path.strip()))
    if not cleaned.lower().endswith(".xlsx"):
        cleaned = f"{cleaned}.xlsx"
    return cleaned
