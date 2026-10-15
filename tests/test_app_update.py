import json
import unittest
from unittest.mock import patch

from services.data_update_service import DEFAULT_DOWNLOAD_URL, check_for_app_update


class AppUpdateTests(unittest.TestCase):
    def check(self, remote, local):
        with patch("services.data_update_service._fetch_bytes", return_value=json.dumps(remote).encode("utf-8")):
            return check_for_app_update(local)

    def test_older_or_equal_remote_does_not_show_update(self):
        for remote, local in (("1.1.3", "1.1.4"), ("1.1.4", "1.1.4"),
                              ("1.1.9", "1.1.10"), ("1.1", "1.1.0"), ("1.0.99", "1.1.0")):
            with self.subTest(remote=remote, local=local):
                self.assertIsNone(self.check({"version": remote}, local))

    def test_only_newer_remote_shows_update(self):
        for remote, local in (("1.1.5", "1.1.4"), ("1.1.10", "1.1.9"),
                              ("1.2.0", "1.1.99"), ("2.0.0", "1.99.99"), ("v1.1.5", "1.1.4")):
            with self.subTest(remote=remote, local=local):
                result = self.check({"version": remote}, local)
                self.assertEqual(result["version"], remote)
                self.assertEqual(result["download_url"], DEFAULT_DOWNLOAD_URL)

    def test_download_url_is_preserved(self):
        remote = {"version": "1.1.5", "download_url": "https://example.org/release"}
        self.assertEqual(self.check(remote, "1.1.4"), remote)

    def test_missing_invalid_or_unrecognized_versions_do_not_show_update(self):
        for remote in (None, [], "1.1.5", {}, {"version": None}, {"version": "latest"}, {"version": "1.1.5-rc1"}):
            with self.subTest(remote=remote):
                self.assertIsNone(self.check(remote, "1.1.4"))
        for local in ("", "unknown", None):
            with self.subTest(local=local):
                self.assertIsNone(self.check({"version": "1.1.5"}, local))

    def test_network_or_json_failure_does_not_interrupt_startup(self):
        with patch("services.data_update_service._fetch_bytes", side_effect=OSError("offline")):
            self.assertIsNone(check_for_app_update("1.1.4"))
        with patch("services.data_update_service._fetch_bytes", return_value=b"invalid json"):
            self.assertIsNone(check_for_app_update("1.1.4"))


if __name__ == "__main__":
    unittest.main()
