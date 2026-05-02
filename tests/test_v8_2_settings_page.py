from __future__ import annotations

import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui.app import (  # noqa: E402
    V8_2_ENABLED_NAV_ITEMS,
    V8_2_SETTINGS_SECTIONS,
    build_rule_config_from_text_fields,
    format_exact_mapping_text,
    format_keyword_mapping_text,
    format_large_piece_companies,
)
from express_app.core.models import ExpressCompanyKeywordRule, ExpressFeeRuleConfig  # noqa: E402


class V82SettingsPageTest(unittest.TestCase):
    def test_settings_sections_match_handoff_structure(self) -> None:
        self.assertEqual(
            list(V8_2_SETTINGS_SECTIONS),
            ["目录配置", "精确映射", "关键词映射", "大件规则"],
        )

    def test_fee_calculation_and_settings_are_enabled_nav_items(self) -> None:
        self.assertEqual(list(V8_2_ENABLED_NAV_ITEMS), ["费用计算", "系统设置"])

    def test_rule_config_can_round_trip_between_text_fields(self) -> None:
        rule_config = ExpressFeeRuleConfig(
            exact_company_map={"申通E物流": "申通", "顺丰速运新3": "顺丰"},
            keyword_company_rules=[
                ExpressCompanyKeywordRule(keyword="顺丰", standard_name="顺丰"),
                ExpressCompanyKeywordRule(keyword="德邦", standard_name="德邦"),
            ],
            large_piece_companies={"顺丰", "德邦"},
            large_piece_threshold_kg=20,
            large_piece_suffix="_大件",
        )

        rebuilt = build_rule_config_from_text_fields(
            exact_mapping_text=format_exact_mapping_text(rule_config),
            keyword_mapping_text=format_keyword_mapping_text(rule_config),
            large_companies_text=format_large_piece_companies(rule_config),
            threshold_text="20",
            suffix_text="_大件",
        )

        self.assertEqual(rebuilt.exact_company_map, rule_config.exact_company_map)
        self.assertEqual(
            [(item.keyword, item.standard_name) for item in rebuilt.keyword_company_rules],
            [("顺丰", "顺丰"), ("德邦", "德邦")],
        )
        self.assertEqual(rebuilt.large_piece_companies, {"顺丰", "德邦"})
        self.assertEqual(rebuilt.large_piece_threshold_kg, 20)
        self.assertEqual(rebuilt.large_piece_suffix, "_大件")

    def test_rule_config_text_validation_reports_empty_threshold(self) -> None:
        with self.assertRaisesRegex(ValueError, "重量阈值必须是数字"):
            build_rule_config_from_text_fields(
                exact_mapping_text="顺丰速运新3=顺丰",
                keyword_mapping_text="顺丰=顺丰",
                large_companies_text="顺丰、德邦",
                threshold_text="",
                suffix_text="_大件",
            )


if __name__ == "__main__":
    unittest.main()
