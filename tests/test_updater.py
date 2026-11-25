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
    _is_allowed_download_url,
    asset_name_for_platform,
    build_posix_helper_script,
    build_windows_helper_script,
    check_for_update,
    decide_apply_prompt,
    download_to,
    fetch_latest_release,
    is_newer,
    parse_version,
    try_startup_auto_update,
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


class TestAllowedDownloadUrl(unittest.TestCase):
    def test_github_release_asset_url_allowed(self) -> None:
        url = (
            "https://github.com/Jeremy-Tarlie/time-tracker/releases/download/v1.7.1/"
            f"{UPDATE_ASSET_WINDOWS}"
        )
        self.assertTrue(_is_allowed_download_url(url))

    def test_release_assets_cdn_allowed(self) -> None:
        url = (
            "https://release-assets.githubusercontent.com/github-production-release-asset/"
            "12345/TimeTracker-windows-x64.exe?sp=r&sv=2026-01-01"
        )
        self.assertTrue(_is_allowed_download_url(url))

    def test_objects_cdn_allowed(self) -> None:
        url = "https://objects.githubusercontent.com/github-production-release-asset/abc"
        self.assertTrue(_is_allowed_download_url(url))

    def test_other_host_rejected(self) -> None:
        self.assertFalse(_is_allowed_download_url(f"https://evil.example/{UPDATE_ASSET_WINDOWS}"))


class TestDownloadToRedirect(unittest.TestCase):
    def test_follows_redirect_to_release_assets_cdn(self) -> None:
        import tempfile
        from email.message import Message
        from pathlib import Path

        source = (
            "https://github.com/Jeremy-Tarlie/time-tracker/releases/download/v1.7.1/"
            f"{UPDATE_ASSET_WINDOWS}"
        )
        cdn = (
            "https://release-assets.githubusercontent.com/github-production-release-asset/"
            f"1/{UPDATE_ASSET_WINDOWS}"
        )
        headers = Message()
        headers["Location"] = cdn
        redirect = urllib.error.HTTPError(
            url=source,
            code=302,
            msg="Found",
            hdrs=headers,
            fp=None,
        )
        ok_response = MagicMock()
        ok_response.status = 200
        ok_response.getcode.return_value = 200
        ok_response.geturl.return_value = cdn
        ok_response.headers = {}
        ok_response.read.side_effect = [b"binary-payload", b""]
        ok_response.__enter__.return_value = ok_response
        ok_response.__exit__.return_value = False

        calls = {"n": 0}

        def _open_side_effect(request: object, *, timeout_s: float) -> object:
            _ = timeout_s
            calls["n"] += 1
            if calls["n"] == 1:
                raise redirect
            return ok_response

        with tempfile.TemporaryDirectory() as tmp:
            dest = str(Path(tmp) / "TimeTracker.exe.new")
            with patch("time_tracker.updater._open_no_redirect", side_effect=_open_side_effect):
                self.assertTrue(download_to(source, dest, timeout_s=1.0, max_bytes=1024))
            self.assertEqual(Path(dest).read_bytes(), b"binary-payload")


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
        self.assertIn(".old.4242", script)
        self.assertIn("update.failed", script)
        self.assertIn(":relaunch", script)
        self.assertIn("Get-Process -Id 4242", script)
        self.assertIn("ping -n 2", script)
        self.assertIn("ping -n 4", script)
        self.assertIn("relaunch", script)
        self.assertNotIn("timeout /t", script)
        self.assertNotIn("| find ", script)
        # Relance même après échec : start après le label relaunch.
        relaunch_idx = script.index(":relaunch")
        self.assertGreater(script.index('start ""', relaunch_idx), relaunch_idx)
        self.assertNotIn("exit /b 1", script)

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
        self.assertIn(".old.1337", script)
        self.assertIn("update.failed", script)
        self.assertIn(">/dev/null 2>&1 &", script)
        self.assertNotIn("exit 1", script)


class TestUpdateFailureMarker(unittest.TestCase):
    def test_has_and_clear_marker(self) -> None:
        import tempfile
        from pathlib import Path

        from time_tracker.updater import (
            clear_update_failure_marker,
            has_update_failure_marker,
            update_failure_marker_path,
        )

        with tempfile.TemporaryDirectory() as tmp:
            exe = str(Path(tmp) / "TimeTracker.exe")
            Path(exe).write_bytes(b"x")
            marker = update_failure_marker_path(exe)
            self.assertFalse(has_update_failure_marker(exe))
            Path(marker).write_text("failed", encoding="utf-8")
            self.assertTrue(has_update_failure_marker(exe))
            clear_update_failure_marker(exe)
            self.assertFalse(has_update_failure_marker(exe))


class TestTryStartupAutoUpdate(unittest.TestCase):
    def test_skips_when_failure_marker_present(self) -> None:
        with (
            patch("time_tracker.updater.is_frozen", return_value=True),
            patch("time_tracker.updater.has_update_failure_marker", return_value=True),
            patch("time_tracker.updater.check_for_update") as check,
            patch("time_tracker.updater.prepare_and_relaunch") as prepare,
        ):
            try_startup_auto_update(current_version="1.7.5")
        check.assert_not_called()
        prepare.assert_not_called()

    def test_up_to_date_removes_stale_new(self) -> None:
        with (
            patch("time_tracker.updater.is_frozen", return_value=True),
            patch("time_tracker.updater.has_update_failure_marker", return_value=False),
            patch(
                "time_tracker.updater.check_for_update",
                return_value=UpdateCheckResult(status="up_to_date"),
            ),
            patch("time_tracker.updater._remove_stale_new_binary") as cleanup,
            patch("time_tracker.updater.apply_update") as apply,
        ):
            try_startup_auto_update(current_version="1.7.7")
        cleanup.assert_called_once_with()
        apply.assert_not_called()


if __name__ == "__main__":
    unittest.main()
