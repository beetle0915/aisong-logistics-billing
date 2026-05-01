"""Small JSON config store for the Tkinter GUI."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from express_app.core.calculator import (
    DEFAULT_OUTPUT_DIR,
    DEFAULT_PRICE_DIR,
    DEFAULT_SPLIT_DIR,
    build_default_rule_config,
    detect_default_sales_file,
)
from express_app.core.models import ExpressCompanyKeywordRule, ExpressFeeRuleConfig


CONFIG_VERSION = 1
APP_DIR = Path(__file__).resolve().parents[2]
LEGACY_CONFIG_DIR = Path.home() / "Library/Application Support/快递费计算工具"


def resolve_default_config_file() -> Path:
    env_path = os.environ.get("EXPRESS_APP_CONFIG_FILE")
    if env_path:
        return Path(env_path).expanduser()
    return APP_DIR / "config.json"


DEFAULT_CONFIG_FILE = resolve_default_config_file()


def _legacy_config_file() -> Path:
    return LEGACY_CONFIG_DIR / "config.json"


@dataclass(frozen=True)
class GuiConfig:
    """Persisted GUI choices."""

    sales_files: list[Path]
    sales_dir: Path
    price_dir: Path
    output_dir: Path
    split_dir: Path
    split_customer_daily_files: bool = True
    generate_customer_history: bool = True
    refresh_all_customers: bool = False
    rule_config: ExpressFeeRuleConfig | None = None


def default_gui_config() -> GuiConfig:
    default_sales = detect_default_sales_file()
    return GuiConfig(
        sales_files=[default_sales],
        sales_dir=default_sales.parent,
        price_dir=DEFAULT_PRICE_DIR,
        output_dir=DEFAULT_OUTPUT_DIR,
        split_dir=DEFAULT_SPLIT_DIR,
        rule_config=build_default_rule_config(),
    )


def _path_from_value(value: Any, default: Path) -> Path:
    if isinstance(value, str) and value.strip():
        return Path(value).expanduser()
    return default


def _bool_from_value(value: Any, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    return default


def _sales_files_from_value(value: Any) -> list[Path]:
    if not isinstance(value, list):
        return []
    paths: list[Path] = []
    for item in value:
        if isinstance(item, str) and item.strip():
            path = Path(item).expanduser()
            try:
                exists = path.exists()
            except OSError:
                exists = False
            if exists:
                paths.append(path)
    return paths


def _string_dict_from_value(value: Any) -> dict[str, str]:
    if not isinstance(value, dict):
        return {}
    result: dict[str, str] = {}
    for raw_name, standard_name in value.items():
        if not isinstance(raw_name, str) or not isinstance(standard_name, str):
            continue
        raw_name = raw_name.strip()
        standard_name = standard_name.strip()
        if raw_name and standard_name:
            result[raw_name] = standard_name
    return result


def _keyword_rules_from_value(value: Any) -> list[ExpressCompanyKeywordRule]:
    if not isinstance(value, list):
        return []
    rules: list[ExpressCompanyKeywordRule] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        keyword = item.get("keyword")
        standard_name = item.get("standard_name")
        if not isinstance(keyword, str) or not isinstance(standard_name, str):
            continue
        keyword = keyword.strip()
        standard_name = standard_name.strip()
        if keyword and standard_name:
            rules.append(
                ExpressCompanyKeywordRule(
                    keyword=keyword,
                    standard_name=standard_name,
                )
            )
    return rules


def _string_set_from_value(value: Any) -> set[str]:
    if not isinstance(value, list):
        return set()
    return {item.strip() for item in value if isinstance(item, str) and item.strip()}


def _float_from_value(value: Any, default: float) -> float:
    if isinstance(value, bool):
        return default
    try:
        parsed = float(value)
    except (TypeError, ValueError):
        return default
    return parsed if parsed > 0 else default


def _rule_config_from_value(value: Any) -> ExpressFeeRuleConfig:
    default = build_default_rule_config()
    if not isinstance(value, dict):
        return default

    exact_map = _string_dict_from_value(value.get("exact_company_map"))
    keyword_rules = _keyword_rules_from_value(value.get("keyword_company_rules"))
    large_piece_companies = _string_set_from_value(value.get("large_piece_companies"))
    suffix_value = value.get("large_piece_suffix")
    large_piece_suffix = (
        suffix_value.strip()
        if isinstance(suffix_value, str) and suffix_value.strip()
        else default.large_piece_suffix
    )

    return ExpressFeeRuleConfig(
        exact_company_map=exact_map or default.exact_company_map,
        keyword_company_rules=keyword_rules or default.keyword_company_rules,
        large_piece_companies=large_piece_companies or default.large_piece_companies,
        large_piece_threshold_kg=_float_from_value(
            value.get("large_piece_threshold_kg"),
            default.large_piece_threshold_kg,
        ),
        large_piece_suffix=large_piece_suffix,
    )


def _rule_config_to_json(rule_config: ExpressFeeRuleConfig) -> dict[str, Any]:
    return {
        "exact_company_map": dict(sorted(rule_config.exact_company_map.items())),
        "keyword_company_rules": [
            {"keyword": item.keyword, "standard_name": item.standard_name}
            for item in rule_config.keyword_company_rules
        ],
        "large_piece_companies": sorted(rule_config.large_piece_companies),
        "large_piece_threshold_kg": rule_config.large_piece_threshold_kg,
        "large_piece_suffix": rule_config.large_piece_suffix,
    }


def load_gui_config(config_file: Path = DEFAULT_CONFIG_FILE) -> GuiConfig:
    """Load GUI config, falling back to defaults on missing or bad JSON."""

    default = default_gui_config()
    if not config_file.exists():
        if config_file.parent.name == "艾松物流计费系统":
            legacy_config_file = _legacy_config_file()
            if legacy_config_file.exists():
                config_file = legacy_config_file
        if not config_file.exists():
            return default

    try:
        raw_data = json.loads(config_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default
    if not isinstance(raw_data, dict):
        return default

    sales_files = _sales_files_from_value(raw_data.get("sales_files"))
    sales_dir = _path_from_value(raw_data.get("sales_dir"), default.sales_dir)
    if not sales_files:
        sales_files = default.sales_files

    return GuiConfig(
        sales_files=sales_files,
        sales_dir=sales_dir,
        price_dir=_path_from_value(raw_data.get("price_dir"), default.price_dir),
        output_dir=_path_from_value(raw_data.get("output_dir"), default.output_dir),
        split_dir=_path_from_value(raw_data.get("split_dir"), default.split_dir),
        split_customer_daily_files=_bool_from_value(
            raw_data.get("split_customer_daily_files"),
            default.split_customer_daily_files,
        ),
        generate_customer_history=_bool_from_value(
            raw_data.get("generate_customer_history"),
            default.generate_customer_history,
        ),
        refresh_all_customers=_bool_from_value(
            raw_data.get("refresh_all_customers"),
            default.refresh_all_customers,
        ),
        rule_config=_rule_config_from_value(raw_data.get("rules")),
    )


def save_gui_config(
    config: GuiConfig,
    config_file: Path = DEFAULT_CONFIG_FILE,
) -> None:
    """Save GUI config. Callers may ignore OSError at the GUI boundary."""

    data = {
        "version": CONFIG_VERSION,
        "sales_dir": str(config.sales_dir),
        "sales_files": [str(path) for path in config.sales_files],
        "price_dir": str(config.price_dir),
        "output_dir": str(config.output_dir),
        "split_dir": str(config.split_dir),
        "split_customer_daily_files": config.split_customer_daily_files,
        "generate_customer_history": config.generate_customer_history,
        "refresh_all_customers": config.refresh_all_customers,
        "rules": _rule_config_to_json(config.rule_config or build_default_rule_config()),
    }
    config_file.parent.mkdir(parents=True, exist_ok=True)
    config_file.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
