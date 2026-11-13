"""Auto-update via GitHub Releases (binaires frozen uniquement)."""

from __future__ import annotations

import json
import os
import platform
import re
import shlex
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Literal, NoReturn, cast
from urllib.parse import urlparse
from urllib.response import addinfourl

from time_tracker.constants import (
    APP_VERSION,
    GITHUB_API_LATEST,
    GITHUB_REPO,
    UPDATE_ASSET_LINUX,
    UPDATE_ASSET_MACOS_ARM64,
    UPDATE_ASSET_WINDOWS,
    UPDATE_CHECK_TIMEOUT_S,
    UPDATE_DOWNLOAD_TIMEOUT_S,
    UPDATE_MAX_DOWNLOAD_BYTES,
)
from time_tracker.logging_setup import get_logger
from time_tracker.paths import get_running_exe, is_frozen

ApplyDecision = Literal["apply_now", "ask_user"]
UpdateCheckStatus = Literal["available", "up_to_date", "error", "unsupported"]

_VERSION_RE = re.compile(r"^v?(\d+(?:\.\d+)*)$", re.IGNORECASE)
_ALLOWED_DOWNLOAD_HOSTS = frozenset({"github.com", "objects.githubusercontent.com"})
_RELEASE_DOWNLOAD_PREFIX = f"/{GITHUB_REPO}/releases/download/"


@dataclass(frozen=True)
class ReleaseAsset:
    name: str
    download_url: str


@dataclass(frozen=True)
class ReleaseInfo:
    tag: str
    assets: tuple[ReleaseAsset, ...]


@dataclass(frozen=True)
class UpdateAvailability:
    current_version: str
    remote_version: str
    download_url: str
    asset_name: str


@dataclass(frozen=True)
class UpdateCheckResult:
    status: UpdateCheckStatus
    availability: UpdateAvailability | None = None


class _NoRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(
        self,
        req: urllib.request.Request,
        fp: object,
        code: int,
        msg: str,
        headers: object,
        newurl: str,
    ) -> urllib.request.Request | None:
        return None


def parse_version(raw: str) -> tuple[int, ...]:
    text = raw.strip()
    match = _VERSION_RE.fullmatch(text)
    if match is None:
        raise ValueError(f"Version invalide: {raw!r}")
    return tuple(int(part) for part in match.group(1).split("."))


def is_newer(remote: str, current: str) -> bool:
    remote_parts = parse_version(remote)
    current_parts = parse_version(current)
    length = max(len(remote_parts), len(current_parts))
    remote_padded = remote_parts + (0,) * (length - len(remote_parts))
    current_padded = current_parts + (0,) * (length - len(current_parts))
    return remote_padded > current_padded


def asset_name_for_platform(system: str, machine: str) -> str | None:
    sys_name = system.lower()
    mach = machine.lower()
    if sys_name == "windows" or sys_name.startswith("win"):
        return UPDATE_ASSET_WINDOWS
    if sys_name == "linux":
        return UPDATE_ASSET_LINUX
    if sys_name == "darwin":
        if mach in {"arm64", "aarch64"}:
            return UPDATE_ASSET_MACOS_ARM64
        return None
    return None


def decide_apply_prompt(*, tracking: bool, writing: bool) -> ApplyDecision:
    if tracking or writing:
        return "ask_user"
    return "apply_now"


def _user_agent() -> str:
    return f"TimeTracker/{APP_VERSION}"


def _is_allowed_download_url(url: str) -> bool:
    parsed = urlparse(url)
    if parsed.scheme != "https":
        return False
    host = (parsed.hostname or "").lower()
    if host not in _ALLOWED_DOWNLOAD_HOSTS:
        return False
    if host == "github.com":
        path = parsed.path or ""
        if not path.startswith(_RELEASE_DOWNLOAD_PREFIX):
            return False
        remainder = path[len(_RELEASE_DOWNLOAD_PREFIX) :]
        if "/" not in remainder:
            return False
        _tag, asset = remainder.split("/", 1)
        if not _tag or "/" in asset or not asset:
            return False
    return True


@contextmanager
def _open_no_redirect(
    request: urllib.request.Request,
    *,
    timeout_s: float,
) -> Iterator[addinfourl]:
    opener = urllib.request.build_opener(_NoRedirectHandler)
    response = cast(addinfourl, opener.open(request, timeout=timeout_s))
    try:
        yield response
    finally:
        response.close()


def fetch_latest_release(
    *,
    timeout_s: float = UPDATE_CHECK_TIMEOUT_S,
) -> ReleaseInfo | None:
    logger = get_logger()
    request = urllib.request.Request(
        GITHUB_API_LATEST,
        headers={
            "User-Agent": _user_agent(),
            "Accept": "application/vnd.github+json",
        },
        method="GET",
    )
    try:
        with _open_no_redirect(request, timeout_s=timeout_s) as response:
            status = getattr(response, "status", None) or response.getcode()
            if status != 200:
                logger.warning("update.check_failed reason=http_status status=%s", status)
                return None
            final_url = getattr(response, "geturl", lambda: GITHUB_API_LATEST)()
            parsed = urlparse(str(final_url))
            if parsed.scheme != "https" or (parsed.hostname or "").lower() != "api.github.com":
                logger.warning("update.check_failed reason=disallowed_final_url")
                return None
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if 300 <= exc.code < 400:
            logger.warning("update.check_failed reason=redirect_disallowed status=%s", exc.code)
        else:
            logger.warning("update.check_failed reason=http_status status=%s", exc.code)
        return None
    except (urllib.error.URLError, TimeoutError, OSError, ValueError, json.JSONDecodeError) as exc:
        logger.warning("update.check_failed reason=%s", type(exc).__name__)
        return None

    if not isinstance(payload, dict):
        logger.warning("update.check_failed reason=invalid_payload")
        return None
    tag = payload.get("tag_name")
    if not isinstance(tag, str) or not tag.strip():
        logger.warning("update.check_failed reason=missing_tag")
        return None

    raw_assets = payload.get("assets")
    if not isinstance(raw_assets, list):
        logger.warning("update.check_failed reason=missing_assets")
        return None

    assets: list[ReleaseAsset] = []
    for item in raw_assets:
        if not isinstance(item, dict):
            continue
        name = item.get("name")
        url = item.get("browser_download_url")
        if isinstance(name, str) and isinstance(url, str) and name and url:
            assets.append(ReleaseAsset(name=name, download_url=url))

    return ReleaseInfo(tag=tag.strip(), assets=tuple(assets))


def check_for_update(
    *,
    current_version: str = APP_VERSION,
    system: str | None = None,
    machine: str | None = None,
    timeout_s: float = UPDATE_CHECK_TIMEOUT_S,
) -> UpdateCheckResult:
    logger = get_logger()
    release = fetch_latest_release(timeout_s=timeout_s)
    if release is None:
        return UpdateCheckResult(status="error")

    try:
        if not is_newer(release.tag, current_version):
            return UpdateCheckResult(status="up_to_date")
    except ValueError:
        logger.warning("update.check_failed reason=invalid_remote_version")
        return UpdateCheckResult(status="error")

    wanted = asset_name_for_platform(
        system if system is not None else platform.system(),
        machine if machine is not None else platform.machine(),
    )
    if wanted is None:
        logger.info("update.skip reason=unsupported_platform")
        return UpdateCheckResult(status="unsupported")

    for asset in release.assets:
        if asset.name == wanted:
            if not _is_allowed_download_url(asset.download_url):
                logger.warning("update.check_failed reason=disallowed_url")
                return UpdateCheckResult(status="error")
            return UpdateCheckResult(
                status="available",
                availability=UpdateAvailability(
                    current_version=current_version,
                    remote_version=release.tag.lstrip("vV"),
                    download_url=asset.download_url,
                    asset_name=asset.name,
                ),
            )

    logger.warning("update.check_failed reason=asset_missing asset=%s", wanted)
    return UpdateCheckResult(status="error")


def download_to(
    url: str,
    dest_path: str,
    *,
    timeout_s: float = UPDATE_DOWNLOAD_TIMEOUT_S,
    max_bytes: int = UPDATE_MAX_DOWNLOAD_BYTES,
    max_redirects: int = 5,
) -> bool:
    logger = get_logger()
    if max_redirects < 0:
        logger.warning("update.download_failed reason=too_many_redirects")
        return False
    if not _is_allowed_download_url(url):
        logger.warning("update.download_failed reason=disallowed_url")
        return False

    parent = os.path.dirname(dest_path)
    if parent:
        os.makedirs(parent, exist_ok=True)

    request = urllib.request.Request(
        url,
        headers={"User-Agent": _user_agent()},
        method="GET",
    )
    tmp_path: str | None = None
    try:
        with _open_no_redirect(request, timeout_s=timeout_s) as response:
            status = getattr(response, "status", None) or response.getcode()
            if status != 200:
                logger.warning("update.download_failed reason=http_status status=%s", status)
                return False
            final_url = str(getattr(response, "geturl", lambda: url)())
            if not _is_allowed_download_url(final_url):
                logger.warning("update.download_failed reason=redirect_disallowed")
                return False
            content_length = (
                response.headers.get("Content-Length") if hasattr(response, "headers") else None
            )
            if content_length is not None:
                try:
                    if int(content_length) > max_bytes:
                        logger.warning("update.download_failed reason=too_large")
                        return False
                except ValueError:
                    pass
            fd, tmp_path = tempfile.mkstemp(
                prefix="timetracker-update-",
                suffix=".partial",
                dir=parent or None,
            )
            size = 0
            with os.fdopen(fd, "wb") as out:
                while True:
                    chunk = response.read(1024 * 256)
                    if not chunk:
                        break
                    size += len(chunk)
                    if size > max_bytes:
                        logger.warning("update.download_failed reason=too_large")
                        return False
                    out.write(chunk)
            if size <= 0:
                logger.warning("update.download_failed reason=empty_file")
                return False
        os.replace(tmp_path, dest_path)
        tmp_path = None
        logger.info("update.downloaded bytes=%s", size)
        return True
    except urllib.error.HTTPError as exc:
        if 300 <= exc.code < 400:
            location = exc.headers.get("Location", "") if exc.headers else ""
            if location and _is_allowed_download_url(location):
                return download_to(
                    location,
                    dest_path,
                    timeout_s=timeout_s,
                    max_bytes=max_bytes,
                    max_redirects=max_redirects - 1,
                )
            logger.warning("update.download_failed reason=redirect_disallowed status=%s", exc.code)
        else:
            logger.warning("update.download_failed reason=http_status status=%s", exc.code)
        return False
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        logger.warning("update.download_failed reason=%s", type(exc).__name__)
        return False
    finally:
        if tmp_path is not None and os.path.isfile(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError as exc:
                logger.warning("update.cleanup_failed reason=%s", type(exc).__name__)


def new_binary_path(exe_path: str | None = None) -> str:
    path = exe_path if exe_path is not None else get_running_exe()
    return f"{path}.new"


def portable_relaunch_args() -> list[str]:
    return ["--portable"] if "--portable" in sys.argv else []


def _batch_quote(path: str) -> str:
    cleaned = path.replace('"', "").replace("%", "").replace("^", "")
    cleaned = cleaned.replace("&", "").replace("|", "").replace("<", "").replace(">", "")
    cleaned = cleaned.replace("(", "").replace(")", "")
    return f'"{cleaned}"'


def build_windows_helper_script(
    *,
    pid: int,
    exe_path: str,
    new_path: str,
    relaunch_args: list[str],
) -> str:
    exe_q = _batch_quote(exe_path)
    new_q = _batch_quote(new_path)
    old_q = _batch_quote(f"{exe_path}.old")
    safe_args = [a for a in relaunch_args if a == "--portable"]
    args = " ".join(f'"{a}"' for a in safe_args)
    launch = f"{exe_q} {args}".rstrip()
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        f":wait\r\n"
        f'tasklist /FI "PID eq {pid}" 2>NUL | find "{pid}" >NUL\r\n'
        "if not errorlevel 1 (\r\n"
        "  timeout /t 1 /nobreak >NUL\r\n"
        "  goto wait\r\n"
        ")\r\n"
        f"if exist {old_q} del /F /Q {old_q}\r\n"
        f"move /Y {exe_q} {old_q}\r\n"
        "if errorlevel 1 exit /b 1\r\n"
        f"move /Y {new_q} {exe_q}\r\n"
        "if errorlevel 1 (\r\n"
        f"  move /Y {old_q} {exe_q}\r\n"
        "  exit /b 1\r\n"
        ")\r\n"
        f'start "" {launch}\r\n'
        f"if exist {old_q} del /F /Q {old_q}\r\n"
        'del /F /Q "%~f0"\r\n'
    )


def build_posix_helper_script(
    *,
    pid: int,
    exe_path: str,
    new_path: str,
    relaunch_args: list[str],
) -> str:
    exe_q = shlex.quote(exe_path)
    new_q = shlex.quote(new_path)
    old_q = shlex.quote(f"{exe_path}.old")
    safe_args = [a for a in relaunch_args if a == "--portable"]
    args_joined = " ".join(shlex.quote(a) for a in safe_args)
    launch = f"{exe_q} {args_joined}".rstrip()
    return (
        "#!/bin/sh\n"
        "set -eu\n"
        f"while kill -0 {pid} 2>/dev/null; do sleep 1; done\n"
        f"rm -f {old_q}\n"
        f"mv {exe_q} {old_q}\n"
        f"if ! mv {new_q} {exe_q}; then\n"
        f"  mv {old_q} {exe_q}\n"
        "  exit 1\n"
        "fi\n"
        f"chmod +x {exe_q}\n"
        f"{launch} >/dev/null 2>&1 &\n"
        f"rm -f {old_q}\n"
        'rm -f -- "$0"\n'
    )


def _spawn_helper(script_path: str) -> None:
    if sys.platform == "win32":
        creationflags = 0
        if hasattr(subprocess, "DETACHED_PROCESS"):
            creationflags |= int(subprocess.DETACHED_PROCESS)
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            creationflags |= int(subprocess.CREATE_NO_WINDOW)
        subprocess.Popen(
            ["cmd.exe", "/c", script_path],
            close_fds=True,
            creationflags=creationflags,
        )
        return

    os.chmod(script_path, 0o755)
    subprocess.Popen(
        ["/bin/sh", script_path],
        close_fds=True,
        start_new_session=True,
    )


def prepare_and_relaunch(new_path: str, *, exe_path: str | None = None) -> NoReturn:
    """Écrit le helper, le lance, puis quitte. Lève OSError si le spawn échoue."""
    logger = get_logger()
    target = exe_path if exe_path is not None else get_running_exe()
    if not os.path.isfile(new_path):
        logger.warning("update.apply_failed reason=missing_new_binary")
        raise FileNotFoundError(new_path)

    relaunch_args = portable_relaunch_args()
    pid = os.getpid()
    if sys.platform == "win32":
        script_path = f"{target}.update.cmd"
        content = build_windows_helper_script(
            pid=pid,
            exe_path=target,
            new_path=new_path,
            relaunch_args=relaunch_args,
        )
        with open(script_path, "w", encoding="utf-8", newline="") as handle:
            handle.write(content)
    else:
        script_path = f"{target}.update.sh"
        content = build_posix_helper_script(
            pid=pid,
            exe_path=target,
            new_path=new_path,
            relaunch_args=relaunch_args,
        )
        with open(script_path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(content)

    logger.info("update.applying")
    # Le lock est libéré à la mort du process (atexit) ; le helper attend le PID.
    try:
        _spawn_helper(script_path)
    except OSError as exc:
        logger.warning("update.apply_failed reason=spawn_failed")
        raise OSError("Impossible de démarrer le helper de mise à jour.") from exc
    sys.exit(0)


def download_update(availability: UpdateAvailability) -> str | None:
    """Télécharge le binaire. Retourne le chemin `.new` ou None."""
    dest = new_binary_path()
    if not download_to(availability.download_url, dest):
        return None
    return dest


def apply_update(availability: UpdateAvailability) -> NoReturn:
    """Télécharge puis applique. À appeler depuis le thread principal."""
    dest = download_update(availability)
    if dest is None:
        raise RuntimeError("Téléchargement de la mise à jour impossible.")
    prepare_and_relaunch(dest)


def try_startup_auto_update(*, current_version: str = APP_VERSION) -> None:
    """Au démarrage frozen : applique silencieusement si une version plus récente existe."""
    if not is_frozen():
        return
    logger = get_logger()
    try:
        result = check_for_update(current_version=current_version)
        if result.status != "available" or result.availability is None:
            return
        logger.info(
            "update.startup_found remote=%s current=%s",
            result.availability.remote_version,
            result.availability.current_version,
        )
        apply_update(result.availability)
    except SystemExit:
        raise
    except Exception as exc:
        logger.warning("update.startup_failed reason=%s", type(exc).__name__)
