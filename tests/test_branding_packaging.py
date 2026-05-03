from __future__ import annotations

import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

import build_macos_app  # noqa: E402
import build_windows_exe  # noqa: E402
from express_app.version import (  # noqa: E402
    APP_DISPLAY_NAME,
    APP_VERSION,
    APP_VERSION_LABEL,
    OUTPUT_VERSION_SUFFIX,
)


class BrandingPackagingTest(unittest.TestCase):
    def test_v8_5_5_branding_metadata(self) -> None:
        self.assertEqual(APP_DISPLAY_NAME, "艾松运费管家")
        self.assertEqual(APP_VERSION, "8.5.5")
        self.assertEqual(APP_VERSION_LABEL, "V8.5.5")
        self.assertEqual(OUTPUT_VERSION_SUFFIX, "v8_5_5")

    def test_packaging_icon_assets_are_wired(self) -> None:
        self.assertEqual(build_macos_app.APP_ICON_FILE, "app_icon.icns")
        self.assertTrue(build_macos_app.APP_ICON_PATH.exists())
        self.assertTrue(build_macos_app.APP_ICON_PATH.name.endswith(".icns"))

        self.assertEqual(build_windows_exe.APP_ICON_PATH.name, "app_icon.ico")
        self.assertTrue(build_windows_exe.APP_ICON_PATH.exists())


if __name__ == "__main__":
    unittest.main()
