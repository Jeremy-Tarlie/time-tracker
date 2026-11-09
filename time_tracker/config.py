from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timedelta

from time_tracker.atomic_io import write_json_atomic
from time_tracker.constants import (
    DEFAULT_SORT,
    ProjectEntry,
    ProjectsConfig,
    SortMode,
)
from time_tracker.hotkeys import default_hotkeys, parse_hotkeys
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
    if value == "recent_desc":
        return "recent_desc"
    return DEFAULT_SORT


def _migrate_legacy_projects(raw: list[str]) -> list[ProjectEntry]:
    base = datetime.now().replace(microsecond=0)
    return [
        {
            "name": name,
            "created": (base + timedelta(seconds=index)).isoformat(timespec="seconds"),
        }
        for index, name in enumerate(raw)
    ]


def _parse_bool(raw: object, *, default: bool) -> bool:
    """Accepte bool / 0 / 1 uniquement ; sinon ``default``."""
    if isinstance(raw, bool):
        return raw
    if isinstance(raw, int) and raw in (0, 1):
        return bool(raw)
    return default


def _fallback_config() -> ProjectsConfig:
    return {
        "projects": _default_projects(),
        "sort_mode": DEFAULT_SORT,
        "excel_file": default_excel_path(),
        "launch_at_startup": False,
        "always_on_top": True,
        "window_geometry": None,
        "hotkeys": default_hotkeys(),
    }


def _quarantine_corrupt(path: str) -> str | None:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = f"{path}.corrupt.{stamp}"
    try:
        os.replace(path, backup)
        return backup
    except OSError:
        return None


@dataclass(frozen=True)
class ConfigLoadResult:
    config: ProjectsConfig
    was_missing: bool = False
    was_corrupt: bool = False
    was_legacy: bool = False
    corrupt_backup_path: str | None = None

    @property
    def should_persist(self) -> bool:
        return self.was_missing or self.was_legacy or self.was_corrupt


def _parse_projects(
    projects_raw: object,
    *,
    allow_empty: bool,
) -> list[ProjectEntry]:
    normalized: list[ProjectEntry] = []
    if isinstance(projects_raw, list):
        for item in projects_raw:
            if isinstance(item, str) and item.strip():
                normalized.append(
                    {
                        "name": item.strip(),
                        "created": datetime.now().isoformat(timespec="seconds"),
                    }
                )
            elif isinstance(item, dict) and "name" in item:
                name = str(item["name"]).strip()
                if not name:
                    continue
                entry: ProjectEntry = {
                    "name": name,
                    "created": str(
                        item.get("created", datetime.now().isoformat(timespec="seconds"))
                    ),
                }
                last_used_raw = item.get("last_used")
                if isinstance(last_used_raw, str) and last_used_raw.strip():
                    entry["last_used"] = last_used_raw.strip()
                normalized.append(entry)
    if normalized:
        return normalized
    if allow_empty and isinstance(projects_raw, list):
        return []
    return _default_projects()


def load_config_detailed(path: str | None = None) -> ConfigLoadResult:
    config_path = path if path is not None else CONFIG_FILE
    if not os.path.exists(config_path):
        return ConfigLoadResult(config=_fallback_config(), was_missing=True)

    try:
        with open(config_path, encoding="utf-8") as handle:
            raw: object = json.load(handle)
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        backup = _quarantine_corrupt(config_path)
        return ConfigLoadResult(
            config=_fallback_config(),
            was_corrupt=True,
            corrupt_backup_path=backup,
        )

    if isinstance(raw, list):
        legacy = [str(item) for item in raw if str(item).strip()]
        return ConfigLoadResult(
            config={
                "projects": (_migrate_legacy_projects(legacy) if legacy else _default_projects()),
                "sort_mode": DEFAULT_SORT,
                "excel_file": default_excel_path(),
                "launch_at_startup": False,
                "always_on_top": True,
                "window_geometry": None,
                "hotkeys": default_hotkeys(),
            },
            was_legacy=True,
        )

    if not isinstance(raw, dict):
        backup = _quarantine_corrupt(config_path)
        return ConfigLoadResult(
            config=_fallback_config(),
            was_corrupt=True,
            corrupt_backup_path=backup,
        )

    projects_raw = raw.get("projects", _default_projects())
    # Fichier valide existant : respecter une liste vide volontaire.
    allow_empty = True
    excel_raw = raw.get("excel_file")
    excel_file = normalize_excel_path(str(excel_raw)) if excel_raw else default_excel_path()
    launch_raw = raw.get("launch_at_startup", False)
    # Défaut True : comportement historique « toujours au premier plan ».
    topmost_raw = raw.get("always_on_top", True)
    geometry_raw = raw.get("window_geometry")
    window_geometry = (
        str(geometry_raw).strip()
        if isinstance(geometry_raw, str) and str(geometry_raw).strip()
        else None
    )
    return ConfigLoadResult(
        config={
            "projects": _parse_projects(projects_raw, allow_empty=allow_empty),
            "sort_mode": _normalize_sort_mode(str(raw.get("sort_mode", DEFAULT_SORT))),
            "excel_file": excel_file,
            "launch_at_startup": _parse_bool(launch_raw, default=False),
            "always_on_top": _parse_bool(topmost_raw, default=True),
            "window_geometry": window_geometry,
            "hotkeys": parse_hotkeys(raw.get("hotkeys")),
        },
    )


def load_config(path: str | None = None) -> ProjectsConfig:
    return load_config_detailed(path).config


def save_config(config: ProjectsConfig, path: str | None = None) -> None:
    config_path = path if path is not None else CONFIG_FILE
    geometry = config.get("window_geometry")
    payload: ProjectsConfig = {
        "projects": config["projects"],
        "sort_mode": config["sort_mode"],
        "excel_file": normalize_excel_path(config["excel_file"]),
        "launch_at_startup": bool(config.get("launch_at_startup", False)),
        "always_on_top": bool(config.get("always_on_top", True)),
        "window_geometry": str(geometry).strip() if geometry else None,
        "hotkeys": parse_hotkeys(config.get("hotkeys")),
    }
    write_json_atomic(config_path, payload)


def sort_projects(projects: list[ProjectEntry], mode: SortMode) -> list[ProjectEntry]:
    if mode == "alpha_asc":
        return sorted(projects, key=lambda entry: entry["name"].casefold())
    if mode == "date_asc":
        return sorted(projects, key=lambda entry: entry["created"])
    if mode == "recent_desc":
        return sorted(
            projects,
            key=lambda entry: entry.get("last_used") or entry["created"],
            reverse=True,
        )
    return sorted(projects, key=lambda entry: entry["created"], reverse=True)


def touch_project_last_used(
    projects: list[ProjectEntry],
    project_name: str,
    *,
    when: datetime | None = None,
) -> list[ProjectEntry]:
    """Met à jour ``last_used`` du projet nommé ; retourne la même liste mutée."""
    stamp = (when or datetime.now()).isoformat(timespec="seconds")
    for entry in projects:
        if entry["name"] == project_name:
            entry["last_used"] = stamp
            break
    return projects


def project_names(projects: list[ProjectEntry]) -> list[str]:
    return [entry["name"] for entry in projects]
