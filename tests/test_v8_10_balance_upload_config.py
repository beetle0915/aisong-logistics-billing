from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui.config_store import GuiConfig, load_gui_config, save_gui_config  # noqa: E402


class V810BalanceUploadConfigTest(unittest.TestCase):
    def test_save_and_load_preserves_balance_upload_endpoint_settings(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            config_path = temp_dir / "config.json"
            sales_file = temp_dir / "sales.xlsx"
            sales_file.write_text("placeholder", encoding="utf-8")

            save_gui_config(
                GuiConfig(
                    sales_files=[sales_file],
                    sales_dir=temp_dir,
                    price_dir=temp_dir / "prices",
                    output_dir=temp_dir / "output",
                    split_dir=temp_dir / "split",
                    balance_upload_url="https://api.example.test/balance",
                    balance_upload_token="secret-token",
                ),
                config_file=config_path,
            )

            loaded = load_gui_config(config_path)

            self.assertEqual(loaded.balance_upload_url, "https://api.example.test/balance")
            self.assertEqual(loaded.balance_upload_token, "secret-token")

    def test_legacy_config_without_balance_upload_settings_uses_empty_defaults(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            config_path = temp_dir / "config.json"
            sales_file = temp_dir / "sales.xlsx"
            sales_file.write_text("placeholder", encoding="utf-8")
            config_path.write_text(
                json.dumps(
                    {
                        "version": 1,
                        "sales_dir": str(temp_dir),
                        "sales_files": [str(sales_file)],
                        "price_dir": str(temp_dir / "prices"),
                        "output_dir": str(temp_dir / "output"),
                        "split_dir": str(temp_dir / "split"),
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            loaded = load_gui_config(config_path)

            self.assertEqual(loaded.balance_upload_url, "")
            self.assertEqual(loaded.balance_upload_token, "")


if __name__ == "__main__":
    unittest.main()
