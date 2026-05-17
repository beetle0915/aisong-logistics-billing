from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.core.models import ExpressCompanyKeywordRule, ExpressFeeRuleConfig  # noqa: E402
from express_app.gui.config_store import GuiConfig, load_gui_config, save_gui_config  # noqa: E402


class V891RuleConfigStoreTest(unittest.TestCase):
    def test_save_and_load_preserves_super_large_piece_rules(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            config_path = temp_dir / "config.json"
            sales_file = temp_dir / "sales.xlsx"
            sales_file.write_text("placeholder", encoding="utf-8")

            rule_config = ExpressFeeRuleConfig(
                exact_company_map={"顺丰速运新3": "顺丰"},
                keyword_company_rules=[
                    ExpressCompanyKeywordRule(keyword="顺丰", standard_name="顺丰"),
                ],
                large_piece_companies={"顺丰", "德邦"},
                large_piece_threshold_kg=20,
                large_piece_suffix="_大件",
                super_large_piece_companies={"顺丰", "德邦"},
                super_large_piece_threshold_kg=60,
                super_large_piece_suffix="_超大件",
            )

            save_gui_config(
                GuiConfig(
                    sales_files=[sales_file],
                    sales_dir=temp_dir,
                    price_dir=temp_dir / "prices",
                    output_dir=temp_dir / "output",
                    split_dir=temp_dir / "split",
                    rule_config=rule_config,
                ),
                config_file=config_path,
            )
            loaded = load_gui_config(config_path)

            self.assertEqual(loaded.rule_config.super_large_piece_companies, {"顺丰", "德邦"})
            self.assertEqual(loaded.rule_config.super_large_piece_threshold_kg, 60)
            self.assertEqual(loaded.rule_config.super_large_piece_suffix, "_超大件")

    def test_legacy_config_without_super_large_piece_rules_uses_defaults(self) -> None:
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
                        "rules": {
                            "exact_company_map": {"顺丰速运新3": "顺丰"},
                            "keyword_company_rules": [
                                {"keyword": "顺丰", "standard_name": "顺丰"},
                            ],
                            "large_piece_companies": ["顺丰", "德邦"],
                            "large_piece_threshold_kg": 20,
                            "large_piece_suffix": "_大件",
                        },
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            loaded = load_gui_config(config_path)

            self.assertEqual(loaded.rule_config.super_large_piece_companies, {"顺丰", "德邦"})
            self.assertEqual(loaded.rule_config.super_large_piece_threshold_kg, 60)
            self.assertEqual(loaded.rule_config.super_large_piece_suffix, "_超大件")


if __name__ == "__main__":
    unittest.main()
