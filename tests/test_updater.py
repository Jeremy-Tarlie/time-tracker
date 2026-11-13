"""Tests unitaires du module d'auto-update (GitHub Releases)."""

from __future__ import annotations

import io
import json
import unittest
import urllib.error
from unittest.mock import MagicMock, patch

from time_tracker.constants import (
    UPDATE_ASSET_LINUX,
    UPDATE_ASSET_MACOS_ARM64,
    UPDATE_ASSET_WINDOWS,
)
from time_tracker.updater import (
    UpdateCheckResult,
    asset_name_for_platform,
    build_posix_helper_script,
    build_windows_helper_script,
    check_for_update,
    decide_apply_prompt,
    fetch_latest_release,
    is_newer,
    parse_version,
)


def _release_json(
    *,
    tag: str = "v1.6.0",
    asset_name: str = UPDATE_ASSET_WINDOWS,
    download_url: str = (
        "https://github.com/Jeremy-Tarlie/time-tracker/releases/download/v1.6.0/"
        f"{UPDATE_ASSET_WINDOWS}"
    ),
) -> bytes:
    payload = {
        "tag_name": tag,
        "assets": [
            {
                "name": asset_name,
                "browser_download_url": download_url,
            }
        ],
    }
    return json.dumps(payload).encode("utf-8")


def _mock_urlopen_response(
    body: bytes,
    status: int = 200,
    *,
    final_url: str = "https://api.github.com/repos/Jeremy-Tarlie/time-tracker/releases/latest",
) -> MagicMock:
    response = MagicMock()
    response.status = status
    response.getcode.return_value = status
    response.read.return_value = body
    response.geturl.return_value = final_url
    response.__enter__.return_value = response
    response.__exit__.return_value = False
    return response


class TestParseVersion(unittest.TestCase):
    def test_parse_plain_semver(self) -> None:
        self.assertEqual(parse_version("1.5.0"), (1, 5, 0))

    def test_parse_strips_lowercase_v_prefix(self) -> None:
        self.assertEqual(parse_version("v1.5.0"), (1, 5, 0))

    def test_parse_strips_uppercase_v_prefix(self) -> None:
        self.assertEqual(parse_version("V2.0"), (2, 0))

    def test_parse_rejects_empty_string(self) -> None:
        with self.assertRaises(ValueError):
            parse_version("")

    def test_parse_rejects_non_numeric(self) -> None:
        with self.assertRaises(ValueError):
            parse_version("abc")

    def test_parse_rejects_non_numeric_segment(self) -> None:
        with self.assertRaises(ValueError):
            parse_version("1.2.x")


class TestIsNewer(unittest.TestCase):
    def test_remote_newer(self) -> None:
        self.assertTrue(is_newer("1.6.0", "1.5.0"))

    def test_versions_equal(self) -> None:
        self.assertFalse(is_newer("1.5.0", "1.5.0"))

    def test_remote_older(self) -> None:
        self.assertFalse(is_newer("1.4.0", "1.5.0"))

    def test_padding_makes_1_5_equal_1_5_0(self) -> None:
        self.assertFalse(is_newer("1.5", "1.5.0"))
        self.assertFalse(is_newer("1.5.0", "1.5"))

    def test_padding_1_5_1_newer_than_1_5(self) -> None:
        self.assertTrue(is_newer("1.5.1", "1.5"))


class TestAssetNameForPlatform(unittest.TestCase):
    def test_windows_returns_windows_asset(self) -> None:
        self.assertEqual(asset_name_for_platform("Windows", "AMD64"), UPDATE_ASSET_WINDOWS)

    def test_win32_returns_windows_asset(self) -> None:
        self.assertEqual(asset_name_for_platform("win32", "x86_64"), UPDATE_ASSET_WINDOWS)

    def test_linux_returns_linux_asset(self) -> None:
        self.assertEqual(asset_name_for_platform("Linux", "x86_64"), UPDATE_ASSET_LINUX)

    def test_darwin_arm64_returns_macos_asset(self) -> None:
        self.assertEqual(
            asset_name_for_platform("Darwin", "arm64"),
            UPDATE_ASSET_MACOS_ARM64,
        )

    def test_darwin_aarch64_returns_macos_asset(self) -> None:
        self.assertEqual(
            asset_name_for_platform("darwin", "aarch64"),
            UPDATE_ASSET_MACOS_ARM64,
        )

    def test_darwin_x86_64_unsupported(self) -> None:
        self.assertIsNone(asset_name_for_platform("Darwin", "x86_64"))

    def test_unknown_system_unsupported(self) -> None:
        self.assertIsNone(asset_name_for_platform("FreeBSD", "amd64"))


class TestDecideApplyPrompt(unittest.TestCase):
    def test_idle_applies_now(self) -> None:
        self.assertEqual(decide_apply_prompt(tracking=False, writing=False), "apply_now")

    def test_tracking_asks_user(self) -> None:
        self.assertEqual(decide_apply_prompt(tracking=True, writing=False), "ask_user")

    def test_writing_asks_user(self) -> None:
        self.assertEqual(decide_apply_prompt(tracking=False, writing=True), "ask_user")

    def test_tracking_and_writing_asks_user(self) -> None:
        self.assertEqual(decide_apply_prompt(tracking=True, writing=True), "ask_user")


class TestFetchLatestRelease(unittest.TestCase):
    def test_returns_release_from_minimal_json(self) -> None:
        body = _release_json(tag="v1.6.0")
        with patch(
            "time_tracker.updater._open_no_redirect",
            return_value=_mock_urlopen_response(body),
        ):
            release = fetch_latest_release(timeout_s=1.0)

        self.assertIsNotNone(release)
        assert release is not None
        self.assertEqual(release.tag, "v1.6.0")
        self.assertEqual(len(release.assets), 1)
        self.assertEqual(release.assets[0].name, UPDATE_ASSET_WINDOWS)
        self.assertTrue(release.assets[0].download_url.startswith("https://github.com/"))

    def test_http_error_returns_none(self) -> None:
        with patch(
            "time_tracker.updater._open_no_redirect",
            side_effect=urllib.error.HTTPError(
                url="https://api.github.com/...",
                code=404,
                msg="Not Found",
                hdrs=None,
                fp=io.BytesIO(b""),
            ),
        ):
            self.assertIsNone(fetch_latest_release(timeout_s=1.0))

    def test_generic_url_error_returns_none(self) -> None:
        with patch(
            "time_tracker.updater._open_no_redirect",
            side_effect=urllib.error.URLError("connection refused"),
        ):
            self.assertIsNone(fetch_latest_release(timeout_s=1.0))


class TestCheckForUpdate(unittest.TestCase):
    def test_newer_remote_with_asset_returns_availability(self) -> None:
        download_url = (
            "https://github.com/Jeremy-Tarlie/time-tracker/releases/download/v1.6.0/"
            f"{UPDATE_ASSET_WINDOWS}"
        )
        release_body = _release_json(
            tag="v1.6.0",
            asset_name=UPDATE_ASSET_WINDOWS,
            download_url=download_url,
        )
        with patch(
            "time_tracker.updater._open_no_redirect",
            return_value=_mock_urlopen_response(release_body),
        ):
            result = check_for_update(
                current_version="1.5.0",
                system="Windows",
                machine="AMD64",
                timeout_s=1.0,
            )

        self.assertIsInstance(result, UpdateCheckResult)
        self.assertEqual(result.status, "available")
        assert result.availability is not None
        self.assertEqual(result.availability.current_version, "1.5.0")
        self.assertEqual(result.availability.remote_version, "1.6.0")
        self.assertEqual(result.availability.asset_name, UPDATE_ASSET_WINDOWS)
        self.assertEqual(result.availability.download_url, download_url)

    def test_same_version_returns_up_to_date(self) -> None:
        release_body = _release_json(tag="v1.5.0")
        with patch(
            "time_tracker.updater._open_no_redirect",
            return_value=_mock_urlopen_response(release_body),
        ):
            result = check_for_update(
                current_version="1.5.0",
                system="Windows",
                machine="AMD64",
                timeout_s=1.0,
            )
        self.assertEqual(result.status, "up_to_date")
        self.assertIsNone(result.availability)

    def test_disallowed_download_url_returns_error(self) -> None:
        release_body = _release_json(
            tag="v1.6.0",
            asset_name=UPDATE_ASSET_WINDOWS,
            download_url=f"http://evil.com/{UPDATE_ASSET_WINDOWS}",
        )
        with patch(
            "time_tracker.updater._open_no_redirect",
            return_value=_mock_urlopen_response(release_body),
        ):
            result = check_for_update(
                current_version="1.5.0",
                system="Windows",
                machine="AMD64",
                timeout_s=1.0,
            )
        self.assertEqual(result.status, "error")
        self.assertIsNone(result.availability)

    def test_other_github_repo_url_returns_error(self) -> None:
        release_body = _release_json(
            tag="v1.6.0",
            asset_name=UPDATE_ASSET_WINDOWS,
            download_url=(
                "https://github.com/other-org/malware/releases/download/v1.6.0/"
                f"{UPDATE_ASSET_WINDOWS}"
            ),
        )
        with patch(
            "time_tracker.updater._open_no_redirect",
            return_value=_mock_urlopen_response(release_body),
        ):
            result = check_for_update(
                current_version="1.5.0",
                system="Windows",
                machine="AMD64",
                timeout_s=1.0,
            )
        self.assertEqual(result.status, "error")


class TestBuildHelperScripts(unittest.TestCase):
    def test_windows_helper_contains_pid_move_start_and_old(self) -> None:
        script = build_windows_helper_script(
            pid=4242,
            exe_path=r"C:\App\TimeTracker.exe",
            new_path=r"C:\App\TimeTracker.exe.new",
            relaunch_args=[],
        )
        self.assertIn("4242", script)
        self.assertIn("move /Y", script)
        self.assertIn('start ""', script)
        self.assertIn(".old", script)

    def test_posix_helper_contains_kill_mv_and_chmod(self) -> None:
        script = build_posix_helper_script(
            pid=1337,
            exe_path="/opt/timetracker/TimeTracker",
            new_path="/opt/timetracker/TimeTracker.new",
            relaunch_args=[],
        )
        self.assertIn("kill -0 1337", script)
        self.assertIn("mv ", script)
        self.assertIn("chmod +x", script)


if __name__ == "__main__":
    unittest.main()
