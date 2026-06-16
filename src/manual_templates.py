"""Manual import templates for cases where Forbes access is unavailable."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .config import (
    CITATION_COLUMNS,
    HISTORY_COLUMNS,
    RAW_DIR,
    TEMPLATES_DIR,
    YearConfig,
    get_year_config,
    ensure_project_dirs,
)


MANUAL_TOP100_TEMPLATE = RAW_DIR / "manual_top100_2025_template.csv"
MANUAL_HISTORY_TEMPLATE = RAW_DIR / "manual_billionaire_wealth_history_template.csv"
MANUAL_CITATIONS_TEMPLATE = RAW_DIR / "manual_source_citations_template.csv"
MANUAL_TOP100_INPUT = RAW_DIR / "manual_top100_2025.csv"
MANUAL_HISTORY_INPUT = RAW_DIR / "manual_billionaire_wealth_history_long.csv"
MANUAL_CITATIONS_INPUT = RAW_DIR / "manual_source_citations.csv"

MANUAL_SOURCE_AUDIT_EXTRA_COLUMNS = [
    "source_mode",
    "entered_by",
    "entered_at",
    "quality_flag",
]

PERSON_EVIDENCE_PACK_COLUMNS = [
    "report_year",
    "rank",
    "forbes_uri",
    "person_name",
    "person_slug",
    "entity_or_asset",
    "wealth_engine_archetype",
    "evidence_category",
    "report_section",
    "citation_key",
    "claim_supported",
    "source_title",
    "source_url",
    "source_file",
    "publisher",
    "author",
    "publication_date",
    "source_as_of_date",
    "claim_year",
    "accessed_at",
    "source_type",
    "reliability_tier",
    "confidence_level",
    "evidence_note",
    "limitations",
]


def _template_paths(config: YearConfig) -> dict[str, Path]:
    if config.legacy_layout:
        return {
            "top100_template": MANUAL_TOP100_TEMPLATE,
            "history_template": MANUAL_HISTORY_TEMPLATE,
            "citations_template": MANUAL_CITATIONS_TEMPLATE,
            "top100_input": MANUAL_TOP100_INPUT,
            "history_input": MANUAL_HISTORY_INPUT,
            "citations_input": MANUAL_CITATIONS_INPUT,
        }
    return {
        "top100_template": TEMPLATES_DIR / f"manual_import_top100_{config.year}.csv",
        "history_template": TEMPLATES_DIR / f"manual_import_wealth_history_{config.year}.csv",
        "citations_template": TEMPLATES_DIR / f"manual_import_source_citations_{config.year}.csv",
        "person_evidence_template": TEMPLATES_DIR / f"manual_import_person_evidence_pack_{config.year}.csv",
        "top100_input": TEMPLATES_DIR / f"manual_import_top100_{config.year}.csv",
        "history_input": TEMPLATES_DIR / f"manual_import_wealth_history_{config.year}.csv",
        "citations_input": TEMPLATES_DIR / f"manual_import_source_citations_{config.year}.csv",
    }


def _manual_input_paths(config: YearConfig, manual_import_dir: str | Path | None = None) -> dict[str, Path]:
    """Return manual input file paths, optionally from a private local directory."""
    if manual_import_dir is None:
        paths = _template_paths(config)
        return {
            "top100_input": paths["top100_input"],
            "history_input": paths["history_input"],
            "citations_input": paths["citations_input"],
        }

    directory = Path(manual_import_dir)
    return {
        "top100_input": directory / f"manual_import_top100_{config.year}.csv",
        "history_input": directory / f"manual_import_wealth_history_{config.year}.csv",
        "citations_input": directory / f"manual_import_source_citations_{config.year}.csv",
    }


def write_manual_import_templates(year: int = 2025) -> dict[str, Path]:
    """Create blank manual-import templates with the required schema."""
    ensure_project_dirs()
    config = get_year_config(year)
    paths = _template_paths(config)
    if not paths["top100_template"].exists():
        pd.DataFrame(columns=config.top100_columns).to_csv(paths["top100_template"], index=False)
    if not paths["history_template"].exists():
        pd.DataFrame(columns=HISTORY_COLUMNS).to_csv(paths["history_template"], index=False)
    if not paths["citations_template"].exists():
        columns = CITATION_COLUMNS if config.legacy_layout else CITATION_COLUMNS + MANUAL_SOURCE_AUDIT_EXTRA_COLUMNS
        pd.DataFrame(columns=columns).to_csv(paths["citations_template"], index=False)
    if not config.legacy_layout and not paths["person_evidence_template"].exists():
        pd.DataFrame(columns=PERSON_EVIDENCE_PACK_COLUMNS).to_csv(paths["person_evidence_template"], index=False)
    return paths


def manual_inputs_available(year: int = 2025, manual_import_dir: str | Path | None = None) -> bool:
    """Return True when both manual input files exist."""
    config = get_year_config(year)
    paths = _manual_input_paths(config, manual_import_dir)
    return paths["top100_input"].exists() and paths["history_input"].exists()


def load_manual_inputs(
    year: int = 2025,
    manual_import_dir: str | Path | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Load user-provided manual top-100 and wealth-history files."""
    config = get_year_config(year)
    paths = _manual_input_paths(config, manual_import_dir)
    if not manual_inputs_available(year, manual_import_dir):
        raise FileNotFoundError(
            f"Manual mode requires {paths['top100_input']} and {paths['history_input']}. "
            "Blank templates have been created; fill them with official annual-list values and citations."
        )
    return pd.read_csv(paths["top100_input"]), pd.read_csv(paths["history_input"])


def load_manual_citations(
    year: int = 2025,
    manual_import_dir: str | Path | None = None,
) -> pd.DataFrame | None:
    """Load optional user-provided manual citation rows."""
    paths = _manual_input_paths(get_year_config(year), manual_import_dir)
    if not paths["citations_input"].exists():
        return None
    return pd.read_csv(paths["citations_input"])
