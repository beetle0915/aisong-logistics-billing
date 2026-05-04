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
from express_app.gui.app import (
    create_startup_lock_overlay,
    remove_startup_lock_overlay,
    run_startup_license_gate,
    should_skip_license_gate,
)  # noqa: E402


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

    def test_startup_license_gate_keeps_main_window_visible_while_prompting(self) -> None:
        class FakeApp:
            def __init__(self) -> None:
                self.withdraw_called = False
                self.destroy_called = False

            def withdraw(self) -> None:
                self.withdraw_called = True

            def destroy(self) -> None:
                self.destroy_called = True

            def update_idletasks(self) -> None:
                pass

        app = FakeApp()

        allowed = run_startup_license_gate(
            app,  # type: ignore[arg-type]
            environ={},
            create_overlay=lambda _app: object(),
            remove_overlay=lambda _overlay: None,
            request_license=lambda _app: True,
        )

        self.assertTrue(allowed)
        self.assertFalse(app.withdraw_called)
        self.assertFalse(app.destroy_called)

    def test_startup_license_gate_closes_app_when_prompt_is_cancelled(self) -> None:
        class FakeApp:
            def __init__(self) -> None:
                self.destroy_called = False

            def destroy(self) -> None:
                self.destroy_called = True

            def update_idletasks(self) -> None:
                pass

        app = FakeApp()

        allowed = run_startup_license_gate(
            app,  # type: ignore[arg-type]
            environ={},
            create_overlay=lambda _app: object(),
            remove_overlay=lambda _overlay: None,
            request_license=lambda _app: False,
        )

        self.assertFalse(allowed)
        self.assertTrue(app.destroy_called)

    def test_startup_license_gate_creates_and_removes_lock_overlay(self) -> None:
        class FakeOverlay:
            def __init__(self) -> None:
                self.destroy_called = False

            def destroy(self) -> None:
                self.destroy_called = True

        class FakeApp:
            def __init__(self) -> None:
                self.overlay = FakeOverlay()
                self.destroy_called = False

            def destroy(self) -> None:
                self.destroy_called = True

            def update_idletasks(self) -> None:
                pass

        app = FakeApp()
        created: list[FakeOverlay] = []

        allowed = run_startup_license_gate(
            app,  # type: ignore[arg-type]
            environ={},
            create_overlay=lambda _app: created.append(app.overlay) or app.overlay,
            remove_overlay=lambda overlay: overlay.destroy(),
            request_license=lambda _app: True,
        )

        self.assertTrue(allowed)
        self.assertEqual(created, [app.overlay])
        self.assertTrue(app.overlay.destroy_called)
        self.assertFalse(app.destroy_called)

    def test_lock_overlay_helpers_mark_app_state(self) -> None:
        class FakeOverlay:
            def __init__(self) -> None:
                self.destroy_called = False

            def destroy(self) -> None:
                self.destroy_called = True

        class FakeApp:
            def __init__(self) -> None:
                self.startup_lock_overlay = None

        app = FakeApp()
        overlay = FakeOverlay()

        result = create_startup_lock_overlay(app, overlay_factory=lambda _app: overlay)
        self.assertIs(result, overlay)
        self.assertIs(app.startup_lock_overlay, overlay)

        remove_startup_lock_overlay(app)
        self.assertTrue(overlay.destroy_called)
        self.assertIsNone(app.startup_lock_overlay)


if __name__ == "__main__":
    unittest.main()
