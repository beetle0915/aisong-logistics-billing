from __future__ import annotations

import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.core.license_key import (  # noqa: E402
    LICENSE_KEY_PROTOCOL_VERSION,
    normalize_license_key,
    generate_license_key,
    verify_license_key,
)
from express_app.gui.app import should_skip_license_gate  # noqa: E402


class V88LicenseKeyTest(unittest.TestCase):
    def test_generates_key_from_protocol_sample_vector(self) -> None:
        self.assertEqual(LICENSE_KEY_PROTOCOL_VERSION, "v1")
        self.assertEqual(generate_license_key(1_800_000_000), "1275-2C7C")
        self.assertEqual(generate_license_key(1_800_000_299), "1275-2C7C")
        self.assertNotEqual(generate_license_key(1_800_000_300), "1275-2C7C")

    def test_verify_accepts_current_previous_and_next_time_windows(self) -> None:
        current_key = generate_license_key(1_800_000_000)
        previous_key = generate_license_key(1_800_000_000 - 300)
        next_key = generate_license_key(1_800_000_000 + 300)

        self.assertTrue(verify_license_key(current_key, 1_800_000_000))
        self.assertTrue(verify_license_key(previous_key, 1_800_000_000))
        self.assertTrue(verify_license_key(next_key, 1_800_000_000))

    def test_verify_rejects_wrong_or_expired_key(self) -> None:
        expired_key = generate_license_key(1_800_000_000 - 600)

        self.assertFalse(verify_license_key("0000-0000", 1_800_000_000))
        self.assertFalse(verify_license_key(expired_key, 1_800_000_000))

    def test_normalize_license_key_accepts_common_copy_formats(self) -> None:
        self.assertEqual(normalize_license_key("12752c7c"), "1275-2C7C")
        self.assertEqual(normalize_license_key(" 1275 2c7c "), "1275-2C7C")
        self.assertEqual(normalize_license_key("1275-2c7c"), "1275-2C7C")
        self.assertEqual(normalize_license_key("abc"), "")

    def test_self_check_can_skip_license_gate(self) -> None:
        self.assertTrue(should_skip_license_gate({"EXPRESS_APP_SELF_CHECK": "1"}))
        self.assertTrue(should_skip_license_gate({"EXPRESS_APP_DISABLE_LICENSE_GATE": "1"}))
        self.assertFalse(should_skip_license_gate({}))


if __name__ == "__main__":
    unittest.main()
