from typing import Literal, TypedDict

SortMode = Literal["alpha_asc", "date_asc", "date_desc"]


class ProjectEntry(TypedDict):
    name: str
    created: str


class ProjectsConfig(TypedDict):
    projects: list[ProjectEntry]
    sort_mode: SortMode
    excel_file: str


DEFAULT_SORT: SortMode = "alpha_asc"
SORT_OPTIONS: list[tuple[str, SortMode]] = [
    ("A → Z", "alpha_asc"),
    ("Date ↑", "date_asc"),
    ("Date ↓", "date_desc"),
]
SORT_LABELS: dict[SortMode, str] = {mode: label for label, mode in SORT_OPTIONS}

APP_VERSION = "1.1.3"
APP_NAME = "Time Tracker"
REGISTRY_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
REGISTRY_NAME = "TimeTracker"
DEFAULT_EXCEL_NAME = "heures_projet.xlsx"
