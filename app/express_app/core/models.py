"""Data models used by the express fee calculation core."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class ExpressCompanyKeywordRule:
    """Keyword-based express company normalization rule."""

    keyword: str
    standard_name: str


@dataclass(frozen=True)
class ExpressFeeRuleConfig:
    """Configurable express fee calculation rules."""

    exact_company_map: dict[str, str] = field(default_factory=dict)
    keyword_company_rules: list[ExpressCompanyKeywordRule] = field(default_factory=list)
    large_piece_companies: set[str] = field(default_factory=set)
    large_piece_threshold_kg: float = 20.0
    large_piece_suffix: str = "_大件"
    super_large_piece_companies: set[str] = field(default_factory=set)
    super_large_piece_threshold_kg: float = 60.0
    super_large_piece_suffix: str = "_超大件"


@dataclass(frozen=True)
class ExpressFeeJobConfig:
    """Input configuration for one express fee calculation job."""

    sales_file: Path
    price_dir: Path
    output_path: Path | None = None
    split_dir: Path | None = None
    round_digits: int | None = 2
    split_customer_daily_files: bool = True
    generate_customer_history: bool = True
    refresh_all_customers: bool = False
    rule_config: ExpressFeeRuleConfig | None = None


@dataclass
class ExpressFeeJobResult:
    """Structured result returned by the calculation core."""

    sales_file: Path
    price_dir: Path
    output_path: Path
    split_dir: Path
    total_rows: int = 0
    success_rows: int = 0
    failed_rows: int = 0
    processing_errors: list[str] = field(default_factory=list)
    split_files: list[Path] = field(default_factory=list)
    split_errors: list[str] = field(default_factory=list)
    history_files: list[Path] = field(default_factory=list)
    history_errors: list[str] = field(default_factory=list)
    logs: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return (
            self.failed_rows == 0
            and not self.processing_errors
            and not self.split_errors
            and not self.history_errors
        )


@dataclass(frozen=True)
class ExpressFeeBatchJobConfig:
    """Input configuration for a batch of sales workbooks."""

    sales_files: list[Path]
    price_dir: Path
    output_dir: Path | None = None
    split_dir: Path | None = None
    round_digits: int | None = 2
    split_customer_daily_files: bool = True
    generate_customer_history: bool = True
    refresh_all_customers: bool = False
    rule_config: ExpressFeeRuleConfig | None = None


@dataclass
class ExpressFeeBatchJobResult:
    """Structured result returned by the batch calculation core."""

    sales_files: list[Path]
    price_dir: Path
    output_dir: Path
    split_dir: Path
    job_results: list[ExpressFeeJobResult] = field(default_factory=list)
    history_files: list[Path] = field(default_factory=list)
    history_errors: list[str] = field(default_factory=list)
    logs: list[str] = field(default_factory=list)

    @property
    def total_rows(self) -> int:
        return sum(result.total_rows for result in self.job_results)

    @property
    def success_rows(self) -> int:
        return sum(result.success_rows for result in self.job_results)

    @property
    def failed_rows(self) -> int:
        return sum(result.failed_rows for result in self.job_results)

    @property
    def ok(self) -> bool:
        return (
            bool(self.job_results)
            and all(result.ok for result in self.job_results)
            and not self.history_errors
        )


@dataclass
class ExpressFeePreflightFileResult:
    """Read-only validation result for one sales workbook."""

    sales_file: Path
    total_rows: int = 0
    success_rows: int = 0
    failed_rows: int = 0
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.failed_rows == 0 and not self.errors


@dataclass
class ExpressFeePreflightResult:
    """Read-only validation result for a batch before writing outputs."""

    sales_files: list[Path]
    price_dir: Path
    file_results: list[ExpressFeePreflightFileResult] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    logs: list[str] = field(default_factory=list)

    @property
    def total_rows(self) -> int:
        return sum(result.total_rows for result in self.file_results)

    @property
    def success_rows(self) -> int:
        return sum(result.success_rows for result in self.file_results)

    @property
    def failed_rows(self) -> int:
        file_failed_rows = sum(result.failed_rows for result in self.file_results)
        if file_failed_rows:
            return file_failed_rows
        return len(self.errors)

    @property
    def ok(self) -> bool:
        return bool(self.file_results) and all(result.ok for result in self.file_results) and not self.errors
