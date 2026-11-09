"""Écriture atomique de fichiers texte / JSON."""

from __future__ import annotations

import json
import os
from typing import TypeVar

T = TypeVar("T")


def write_text_atomic(path: str, content: str, *, encoding: str = "utf-8") -> None:
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp_path = f"{path}.tmp"
    try:
        with open(tmp_path, "w", encoding=encoding) as handle:
            handle.write(content)
        os.replace(tmp_path, path)
    except OSError:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise


def write_json_atomic(path: str, payload: object) -> None:
    write_text_atomic(
        path,
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
    )
