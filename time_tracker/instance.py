"""Verrou d'instance unique (fichier exclusif)."""

from __future__ import annotations

import atexit
import importlib
import os
import sys
from collections.abc import Callable
from types import TracebackType
from typing import IO, TextIO, cast

from time_tracker.paths import INSTANCE_LOCK_FILE

_lock_handle: TextIO | None = None


class InstanceLockError(RuntimeError):
    """Une autre instance détient déjà le verrou."""


def _lock_windows(handle: IO[str]) -> None:
    msvcrt_mod = importlib.import_module("msvcrt")
    locking = cast(Callable[[int, int, int], None], msvcrt_mod.locking)
    locking(handle.fileno(), int(msvcrt_mod.LK_NBLCK), 1)


def _unlock_windows(handle: IO[str]) -> None:
    msvcrt_mod = importlib.import_module("msvcrt")
    locking = cast(Callable[[int, int, int], None], msvcrt_mod.locking)
    try:
        handle.seek(0)
        locking(handle.fileno(), int(msvcrt_mod.LK_UNLCK), 1)
    except OSError:
        pass


def _lock_posix(handle: IO[str]) -> None:
    fcntl_mod = importlib.import_module("fcntl")
    flock = cast(Callable[[int, int], None], fcntl_mod.flock)
    flock(handle.fileno(), int(fcntl_mod.LOCK_EX) | int(fcntl_mod.LOCK_NB))


def _unlock_posix(handle: IO[str]) -> None:
    fcntl_mod = importlib.import_module("fcntl")
    flock = cast(Callable[[int, int], None], fcntl_mod.flock)
    try:
        flock(handle.fileno(), int(fcntl_mod.LOCK_UN))
    except OSError:
        pass


def acquire_instance_lock(path: str = INSTANCE_LOCK_FILE) -> TextIO:
    global _lock_handle
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    handle = open(path, "a+", encoding="utf-8")
    try:
        handle.seek(0)
        handle.truncate()
        handle.write(str(os.getpid()))
        handle.flush()
        if sys.platform == "win32":
            _lock_windows(handle)
        else:
            _lock_posix(handle)
    except OSError as exc:
        handle.close()
        raise InstanceLockError("Time Tracker est déjà en cours d'exécution.") from exc
    _lock_handle = handle
    atexit.register(release_instance_lock)
    return handle


def release_instance_lock() -> None:
    global _lock_handle
    handle = _lock_handle
    if handle is None:
        return
    try:
        if sys.platform == "win32":
            _unlock_windows(handle)
        else:
            _unlock_posix(handle)
    finally:
        try:
            handle.close()
        except OSError:
            pass
        _lock_handle = None


class InstanceLock:
    def __init__(self, path: str = INSTANCE_LOCK_FILE) -> None:
        self.path = path
        self._handle: TextIO | None = None

    def __enter__(self) -> InstanceLock:
        self._handle = acquire_instance_lock(self.path)
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        release_instance_lock()


def ensure_single_instance() -> bool:
    """Acquiert le verrou. Retourne False si une autre instance tourne déjà."""
    try:
        acquire_instance_lock()
        return True
    except InstanceLockError:
        return False


def exit_if_already_running() -> None:
    if not ensure_single_instance():
        print("Time Tracker est déjà en cours d'exécution.", file=sys.stderr)
        sys.exit(0)
