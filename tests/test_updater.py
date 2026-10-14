from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

APP_DIR = Path(__file__).resolve().parents[1] / "outputs" / "c6_streamlit_app"
sys.path.insert(0, str(APP_DIR))

from updater.checksum import read_expected_sha256, verify_sha256  # noqa: E402
from updater.github_release import ReleaseInfo, get_latest_release  # noqa: E402
from updater.update_manager import UpdateState, check_for_update  # noqa: E402
from updater_app import install_update  # noqa: E402


class UpdaterTests(unittest.TestCase):
    def test_release_requires_all_verified_assets(self):
        payload = {
            "tag_name": "v2.7.3",
            "assets": [
                {"name": "AutoC6Bank.exe", "browser_download_url": "https://example/exe"},
                {"name": "AutoC6Bank.exe.sha256", "browser_download_url": "https://example/exe-sha"},
                {"name": "AutoC6BankUpdater.exe", "browser_download_url": "https://example/updater"},
            ],
        }

        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            def read(self):
                return json.dumps(payload).encode("utf-8")

        self.assertIsNone(get_latest_release(opener=lambda *_args, **_kwargs: Response()))

    def test_semver_comparison_is_not_lexical(self):
        release = ReleaseInfo("v1.10.0", "1.10.0", "", "https://x/exe", "https://x/sum", "https://x/updater")
        with mock.patch("updater.update_manager.get_latest_release", return_value=release):
            result = check_for_update(current_version="1.9.9")
        self.assertEqual(result.state, UpdateState.AVAILABLE)

    def test_same_version_has_no_update(self):
        release = ReleaseInfo("v1.0.0", "1.0.0", "", "https://x/exe", "https://x/sum", "https://x/updater")
        with mock.patch("updater.update_manager.get_latest_release", return_value=release):
            self.assertEqual(check_for_update(current_version="1.0.0").state, UpdateState.IDLE)

    def test_github_failure_does_not_block_app(self):
        with mock.patch("updater.update_manager.get_latest_release", side_effect=OSError("offline")):
            self.assertEqual(check_for_update().state, UpdateState.ERROR)

    def test_checksum_valid_and_invalid(self):
        with tempfile.TemporaryDirectory() as folder:
            payload = Path(folder) / "app.exe"
            checksum = Path(folder) / "app.exe.sha256"
            payload.write_bytes(b"conteudo")
            digest = hashlib.sha256(b"conteudo").hexdigest()
            checksum.write_text(f"{digest}  AutoC6Bank.exe\n", encoding="utf-8")
            self.assertTrue(verify_sha256(payload, read_expected_sha256(checksum)))
            self.assertFalse(verify_sha256(payload, "0" * 64))

    def test_rollback_restores_old_executable_if_replacement_fails(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            target, new, backup = root / "AutoC6Bank.exe", root / "AutoC6Bank.new.exe", root / "AutoC6Bank.old.exe"
            target.write_bytes(b"old")
            new.write_bytes(b"new")
            with mock.patch("updater_app._wait_for_exit", return_value=True), mock.patch(
                "updater_app.subprocess.Popen", side_effect=OSError("nao inicia")
            ):
                with self.assertRaises(OSError):
                    install_update(target, new, backup, 123)
            self.assertEqual(target.read_bytes(), b"old")


if __name__ == "__main__":
    unittest.main()
