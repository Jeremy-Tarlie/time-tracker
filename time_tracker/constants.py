from typing import Literal, TypedDict

SortMode = Literal["alpha_asc", "date_asc", "date_desc", "recent_desc"]
HotkeyAction = Literal[
    "toggle",
    "stop",
    "settings",
    "show",
    "today",
    "period",
    "manual",
]


class HotkeysConfig(TypedDict):
    toggle: str
    stop: str
    settings: str
    show: str
    today: str
    period: str
    manual: str


class _ProjectEntryRequired(TypedDict):
    name: str
    created: str


class ProjectEntry(_ProjectEntryRequired, total=False):
    last_used: str


class ProjectsConfig(TypedDict):
    projects: list[ProjectEntry]
    sort_mode: SortMode
    excel_file: str
    launch_at_startup: bool
    always_on_top: bool
    window_geometry: str | None
    hotkeys: HotkeysConfig


class PendingSession(TypedDict):
    project: str
    start: str
    end: str
    notes: str


class ActiveSession(TypedDict):
    project: str
    start: str
    accumulated_seconds: float
    paused: bool
    resume_at: str | None


class SessionRow(TypedDict):
    row_index: int
    date: str
    project: str
    start: str
    end: str
    duration: str
    notes: str


DEFAULT_SORT: SortMode = "alpha_asc"
SORT_OPTIONS: list[tuple[str, SortMode]] = [
    ("A → Z", "alpha_asc"),
    ("Date ↑", "date_asc"),
    ("Date ↓", "date_desc"),
    ("Récents", "recent_desc"),
]
SORT_LABELS: dict[SortMode, str] = {mode: label for label, mode in SORT_OPTIONS}

APP_VERSION = "1.6.0"
APP_NAME = "Time Tracker"
REGISTRY_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
REGISTRY_NAME = "TimeTracker"
DEFAULT_EXCEL_NAME = "heures_projet.xlsx"
PENDING_SESSIONS_NAME = "pending_sessions.json"
ACTIVE_SESSION_NAME = "active_session.json"
INSTANCE_LOCK_NAME = "timetracker.lock"
PORTABLE_FLAG_NAME = "portable.flag"

GITHUB_REPO = "Jeremy-Tarlie/time-tracker"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
UPDATE_ASSET_WINDOWS = "TimeTracker-windows-x64.exe"
UPDATE_ASSET_LINUX = "TimeTracker-linux-x64"
UPDATE_ASSET_MACOS_ARM64 = "TimeTracker-macos-arm64"
UPDATE_CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000
UPDATE_CHECK_TIMEOUT_S = 15
UPDATE_DOWNLOAD_TIMEOUT_S = 180
UPDATE_MAX_DOWNLOAD_BYTES = 200 * 1024 * 1024
