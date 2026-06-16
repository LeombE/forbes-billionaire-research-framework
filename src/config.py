"""Project configuration and constants."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]
CONFIG_DIR = BASE_DIR / "config"
DATA_DIR = BASE_DIR / "data"
RAW_DIR = DATA_DIR / "raw"
INTERIM_DIR = DATA_DIR / "interim"
PROCESSED_DIR = DATA_DIR / "processed"
TEMPLATES_DIR = BASE_DIR / "templates"
REPORTS_DIR = BASE_DIR / "reports"
CHARTS_DIR = REPORTS_DIR / "charts"
PEOPLE_REPORTS_DIR = REPORTS_DIR / "people"
TESTS_DIR = BASE_DIR / "tests"

FORBES_ROBOTS_URL = "https://www.forbes.com/robots.txt"
FORBES_TERMS_URL = "https://www.forbes.com/terms/"
FORBES_2025_TOP200_ARTICLE_URL = (
    "https://www.forbes.com/sites/chasewithorn/2025/04/01/"
    "forbes-worlds-billionaires-list-2025-the-top-200/"
)
FORBES_ANNUAL_API_TEMPLATE = (
    "https://www.forbes.com/forbesapi/person/billionaires/"
    "{year}/position/true.json?limit={limit}"
)
FORBES_PROFILE_TEMPLATE = "https://www.forbes.com/profile/{uri}/"

HISTORY_START_YEAR = 2001
CANONICAL_YEAR = 2025
DEFAULT_TARGET_YEAR = 2026
HISTORY_YEARS = list(range(HISTORY_START_YEAR, CANONICAL_YEAR + 1))
FORBES_API_LIMIT = 4000

REQUEST_TIMEOUT_SECONDS = 45
REQUEST_DELAY_SECONDS = float(os.getenv("FORBES_REQUEST_DELAY_SECONDS", "1.5"))
USER_AGENT = os.getenv(
    "FORBES_RESEARCH_USER_AGENT",
    "Mozilla/5.0 (compatible; local-forbes-research-project/1.0; educational research)",
)

TOP100_COLUMNS = [
    "rank_2025",
    "position_2025",
    "forbes_uri",
    "name",
    "net_worth_2025_usd_b",
    "age_2025",
    "country_or_territory",
    "country_of_citizenship",
    "source_of_wealth",
    "industry",
    "primary_company_or_asset",
    "self_made_or_inherited_if_available",
    "forbes_profile_url",
    "canonical_source_url",
    "notes",
]

TOP100_MULTIYEAR_COLUMNS = [
    "year",
    "rank",
    "position",
    "forbes_uri",
    "name",
    "name_slug",
    "net_worth_usd_b",
    "age",
    "country_or_territory",
    "country_of_citizenship",
    "source_of_wealth",
    "industry",
    "primary_company_or_asset",
    "self_made_or_inherited_if_available",
    "forbes_profile_url",
    "canonical_annual_list_source_url",
    "canonical_source_url",
    "notes",
    "data_status",
    "source_quality",
]

HISTORY_COLUMNS = [
    "forbes_uri",
    "name",
    "year",
    "rank",
    "position",
    "net_worth_usd_b",
    "age",
    "country_or_territory",
    "source_of_wealth",
    "industry",
    "primary_company_or_asset",
    "forbes_profile_url",
    "source_url",
]

GROWTH_COLUMNS = [
    "forbes_uri",
    "name",
    "rank_2025",
    "first_year_observed",
    "first_net_worth_usd_b",
    "years_observed",
    "observation_span_years",
    "wealth_multiple_first_to_2025",
    "CAGR_nominal",
    "log_linear_growth_slope",
    "exponential_fit_r2",
    "estimated_doubling_time_years",
    "peak_net_worth_usd_b",
    "max_one_year_gain_usd_b",
    "max_one_year_loss_usd_b",
    "volatility_of_annual_growth",
    "largest_drawdown_pct",
    "wealth_engine_category",
    "secondary_wealth_engines",
    "classification_confidence",
    "public_equity_dependency_flag",
    "key_asset_or_company",
    "evidence_summary",
    "source_ids",
    "data_completeness_score",
]

GROWTH_MULTIYEAR_COLUMNS = [
    "year",
    "rank",
    "forbes_uri",
    "name",
    "first_year_observed",
    "first_net_worth_usd_b",
    "years_observed",
    "observation_span_years",
    "wealth_multiple_first_to_target_year",
    "CAGR_nominal",
    "log_linear_growth_slope",
    "exponential_fit_r2",
    "estimated_doubling_time_years",
    "peak_net_worth_usd_b",
    "max_one_year_gain_usd_b",
    "max_one_year_loss_usd_b",
    "volatility_of_annual_growth",
    "largest_drawdown_pct",
    "wealth_engine_category",
    "secondary_wealth_engines",
    "classification_confidence",
    "public_equity_dependency_flag",
    "key_asset_or_company",
    "evidence_summary",
    "source_ids",
    "data_completeness_score",
]

CITATION_COLUMNS = [
    "source_id",
    "citation_scope",
    "person_id",
    "forbes_uri",
    "person_name",
    "name",
    "table_name",
    "row_key",
    "year",
    "field_name",
    "claim_supported",
    "source_title",
    "source_url",
    "publisher",
    "author",
    "publication_date",
    "accessed_at",
    "accessed_at_utc",
    "source_type",
    "evidence_note",
    "reliability_tier",
    "limitations",
    "derived_from",
    "source_file",
    "collection_method",
    "notes",
]

PERSON_QUALITY_COLUMNS = [
    "forbes_uri",
    "name",
    "rank_2025",
    "top100_required_fields_complete",
    "years_observed",
    "has_growth_metrics",
    "has_required_citations",
    "docx_report_ready",
    "data_quality_score",
    "missing_or_limited_evidence",
]

PERSON_QUALITY_MULTIYEAR_COLUMNS = [
    "year",
    "forbes_uri",
    "name",
    "rank",
    "top100_required_fields_complete",
    "years_observed",
    "has_growth_metrics",
    "has_required_citations",
    "docx_report_ready",
    "data_quality_score",
    "missing_or_limited_evidence",
]

REPORT_SECTIONS = [
    "Executive thesis",
    "Wealth equation",
    "Ownership and asset map",
    "Wealth history and CAGR",
    "Exponential-fit result and interpretation",
    "Business empire timeline",
    "Core operating companies/assets",
    "Industry structure",
    "Financial statement connection",
    "Moat and competitive advantage",
    "Capital allocation pattern",
    "Country/economic/regulatory context",
    "Key strategic decisions",
    "Risk factors",
    "Comparable billionaires",
    "Evidence table",
    "Data limitations",
]

WEALTH_ENGINE_CATEGORIES = [
    "Founder/operator public equity",
    "Founder/operator private company",
    "Early employee/executive equity",
    "Investor/capital allocator",
    "Inherited/family-controlled business",
    "Luxury/retail brand ownership",
    "Technology/platform monopoly/network effects",
    "Real estate/land/infrastructure",
    "Commodities/energy/resources",
    "Diversified holding company",
    "Other/unclear",
]


@dataclass(frozen=True)
class YearConfig:
    """Resolved paths and schema names for one canonical annual list year."""

    year: int
    canonical_source: str = "Forbes World's Billionaires annual list"

    @property
    def legacy_layout(self) -> bool:
        return self.year == 2025

    @property
    def history_years(self) -> list[int]:
        return list(range(HISTORY_START_YEAR, self.year + 1))

    @property
    def raw_forbes_dir(self) -> Path:
        return RAW_DIR if self.legacy_layout else RAW_DIR / "forbes" / str(self.year)

    @property
    def interim_dir(self) -> Path:
        return INTERIM_DIR if self.legacy_layout else INTERIM_DIR / str(self.year)

    @property
    def processed_dir(self) -> Path:
        return PROCESSED_DIR if self.legacy_layout else PROCESSED_DIR / str(self.year)

    @property
    def annual_reports_dir(self) -> Path:
        return REPORTS_DIR if self.legacy_layout else REPORTS_DIR / "annual" / str(self.year)

    @property
    def charts_dir(self) -> Path:
        return CHARTS_DIR if self.legacy_layout else CHARTS_DIR / str(self.year)

    @property
    def people_reports_dir(self) -> Path:
        return PEOPLE_REPORTS_DIR if self.legacy_layout else PEOPLE_REPORTS_DIR / str(self.year)

    @property
    def rank_col(self) -> str:
        return f"rank_{self.year}" if self.legacy_layout else "rank"

    @property
    def position_col(self) -> str:
        return f"position_{self.year}" if self.legacy_layout else "position"

    @property
    def net_worth_col(self) -> str:
        return f"net_worth_{self.year}_usd_b" if self.legacy_layout else "net_worth_usd_b"

    @property
    def age_col(self) -> str:
        return f"age_{self.year}" if self.legacy_layout else "age"

    @property
    def multiple_col(self) -> str:
        return f"wealth_multiple_first_to_{self.year}" if self.legacy_layout else "wealth_multiple_first_to_target_year"

    @property
    def top100_columns(self) -> list[str]:
        return TOP100_COLUMNS if self.legacy_layout else TOP100_MULTIYEAR_COLUMNS

    @property
    def growth_columns(self) -> list[str]:
        return GROWTH_COLUMNS if self.legacy_layout else GROWTH_MULTIYEAR_COLUMNS

    @property
    def person_quality_columns(self) -> list[str]:
        return PERSON_QUALITY_COLUMNS if self.legacy_layout else PERSON_QUALITY_MULTIYEAR_COLUMNS

    @property
    def top100_filename(self) -> str:
        return f"top100_{self.year}.csv"

    @property
    def history_filename(self) -> str:
        return (
            "billionaire_wealth_history_long.csv"
            if self.legacy_layout
            else f"billionaire_wealth_history_long_{self.year}.csv"
        )

    @property
    def metrics_filename(self) -> str:
        return (
            "billionaire_growth_metrics.csv"
            if self.legacy_layout
            else f"billionaire_growth_metrics_{self.year}.csv"
        )

    @property
    def citations_filename(self) -> str:
        return "source_citations.csv" if self.legacy_layout else f"source_citations_{self.year}.csv"

    @property
    def person_quality_filename(self) -> str:
        return "person_data_quality_scores.csv" if self.legacy_layout else f"person_data_quality_scores_{self.year}.csv"

    @property
    def data_quality_checks_filename(self) -> str:
        return "data_quality_checks.csv" if self.legacy_layout else f"data_quality_checks_{self.year}.csv"

    @property
    def missing_summary_filename(self) -> str:
        return "missing_field_summary.csv" if self.legacy_layout else f"missing_field_summary_{self.year}.csv"

    @property
    def realtime_comparison_filename(self) -> str:
        return f"forbes_realtime_comparison_june_{self.year}.csv"

    @property
    def workbook_path(self) -> Path:
        return BASE_DIR / f"Forbes_top100_{self.year}_analysis.xlsx"

    @property
    def top100_path(self) -> Path:
        return self.processed_dir / self.top100_filename

    @property
    def history_path(self) -> Path:
        return self.processed_dir / self.history_filename

    @property
    def metrics_path(self) -> Path:
        return self.processed_dir / self.metrics_filename

    @property
    def citations_path(self) -> Path:
        return self.processed_dir / self.citations_filename

    @property
    def person_quality_path(self) -> Path:
        return self.processed_dir / self.person_quality_filename

    @property
    def top100_sheet(self) -> str:
        return f"Top100_{self.year}"

    @property
    def wealth_history_sheet(self) -> str:
        return "Wealth_History_Long" if self.legacy_layout else f"Wealth_History_Long_{self.year}"

    @property
    def growth_metrics_sheet(self) -> str:
        return "Growth_Metrics" if self.legacy_layout else f"Growth_Metrics_{self.year}"

    @property
    def industry_summary_sheet(self) -> str:
        return "Industry_Summary" if self.legacy_layout else f"Industry_Summary_{self.year}"

    @property
    def country_summary_sheet(self) -> str:
        return "Country_Summary" if self.legacy_layout else f"Country_Summary_{self.year}"

    @property
    def wealth_engine_summary_sheet(self) -> str:
        return "Wealth_Engine_Summary" if self.legacy_layout else f"Wealth_Engine_Summary_{self.year}"

    @property
    def source_citations_sheet(self) -> str:
        return "Source_Citations" if self.legacy_layout else f"Source_Citations_{self.year}"

    @property
    def evidence_registry_sheet(self) -> str:
        return f"Enriched_Evidence_Registry_{self.year}"

    @property
    def data_quality_sheet(self) -> str:
        return "Data_Quality" if self.legacy_layout else f"Data_Quality_{self.year}"

    @property
    def methodology_sheet(self) -> str:
        return "Methodology" if self.legacy_layout else f"Methodology_{self.year}"

    @property
    def source_file_for_year(self) -> str:
        if self.legacy_layout:
            return f"data/raw/forbes_billionaires_{self.year}.json"
        return f"data/raw/forbes/{self.year}/forbes_billionaires_{self.year}.json"


def get_year_config(year: int = DEFAULT_TARGET_YEAR) -> YearConfig:
    """Return the project configuration for one annual canonical year."""
    return YearConfig(int(year))


def annual_api_url(year: int, limit: int = FORBES_API_LIMIT) -> str:
    """Return the Forbes annual-list API URL for a year."""
    return FORBES_ANNUAL_API_TEMPLATE.format(year=year, limit=limit)


def profile_url(uri: str | None) -> str:
    """Return the canonical Forbes profile URL for a Forbes profile URI."""
    if not uri:
        return ""
    return FORBES_PROFILE_TEMPLATE.format(uri=str(uri).strip("/"))


def ensure_project_dirs() -> None:
    """Create the output directory tree."""
    for directory in [
        CONFIG_DIR,
        DATA_DIR,
        RAW_DIR,
        INTERIM_DIR,
        PROCESSED_DIR,
        TEMPLATES_DIR,
        REPORTS_DIR,
        CHARTS_DIR,
        PEOPLE_REPORTS_DIR,
        TESTS_DIR,
    ]:
        directory.mkdir(parents=True, exist_ok=True)


def ensure_year_dirs(config: YearConfig) -> None:
    """Create the year-aware directory tree without removing legacy folders."""
    ensure_project_dirs()
    for directory in [
        config.raw_forbes_dir,
        config.interim_dir,
        config.processed_dir,
        config.annual_reports_dir,
        config.charts_dir,
        config.people_reports_dir,
    ]:
        directory.mkdir(parents=True, exist_ok=True)
