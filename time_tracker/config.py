import json
import os
from datetime import datetime, timedelta

from time_tracker.constants import (
    DEFAULT_SORT,
    ProjectEntry,
    ProjectsConfig,
    SortMode,
)
from time_tracker.paths import CONFIG_FILE, default_excel_path, normalize_excel_path


def _default_projects() -> list[ProjectEntry]:
    now = datetime.now().isoformat(timespec="seconds")
    return [
        {"name": "Projet 1", "created": now},
        {"name": "Projet 2", "created": now},
        {"name": "Projet 3", "created": now},
    ]


def _normalize_sort_mode(value: str) -> SortMode:
    if value == "alpha_asc":
        return "alpha_asc"
    if value == "date_asc":
        return "date_asc"
    if value == "date_desc":
        return "date_desc"
    return DEFAULT_SORT


def _migrate_legacy_projects(raw: list[str]) -> list[ProjectEntry]:
    base = datetime.now().replace(microsecond=0)
    return [
        {"name": name, "created": (base + timedelta(seconds=index)).isoformat(timespec="seconds")}
        for index, name in enumerate(raw)
    ]


def load_config() -> ProjectsConfig:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as handle:
            raw: object = json.load(handle)
        if isinstance(raw, list):
            legacy = [str(item) for item in raw]
            return {
                "projects": _migrate_legacy_projects(legacy),
                "sort_mode": DEFAULT_SORT,
                "excel_file": default_excel_path(),
            }
        if not isinstance(raw, dict):
            return {
                "projects": _default_projects(),
                "sort_mode": DEFAULT_SORT,
                "excel_file": default_excel_path(),
            }
        projects_raw = raw.get("projects", _default_projects())
        normalized: list[ProjectEntry] = []
        if isinstance(projects_raw, list):
            for item in projects_raw:
                if isinstance(item, str):
                    normalized.append({
                        "name": item,
                        "created": datetime.now().isoformat(timespec="seconds"),
                    })
                elif isinstance(item, dict):
                    normalized.append({
                        "name": str(item["name"]),
                        "created": str(
                            item.get("created", datetime.now().isoformat(timespec="seconds"))
                        ),
                    })
        excel_raw = raw.get("excel_file")
        excel_file = (
            normalize_excel_path(str(excel_raw)) if excel_raw else default_excel_path()
        )
        return {
            "projects": normalized if normalized else _default_projects(),
            "sort_mode": _normalize_sort_mode(str(raw.get("sort_mode", DEFAULT_SORT))),
            "excel_file": excel_file,
        }
    return {
        "projects": _default_projects(),
        "sort_mode": DEFAULT_SORT,
        "excel_file": default_excel_path(),
    }


def save_config(config: ProjectsConfig) -> None:
    payload: ProjectsConfig = {
        "projects": config["projects"],
        "sort_mode": config["sort_mode"],
        "excel_file": normalize_excel_path(config["excel_file"]),
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def sort_projects(projects: list[ProjectEntry], mode: SortMode) -> list[ProjectEntry]:
    if mode == "alpha_asc":
        return sorted(projects, key=lambda entry: entry["name"].casefold())
    if mode == "date_asc":
        return sorted(projects, key=lambda entry: entry["created"])
    return sorted(projects, key=lambda entry: entry["created"], reverse=True)


def project_names(projects: list[ProjectEntry]) -> list[str]:
    return [entry["name"] for entry in projects]
