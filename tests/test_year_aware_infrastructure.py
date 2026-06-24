from __future__ import annotations

import inspect
import math
import subprocess
from pathlib import Path

import pandas as pd

from src.archetype_registry import build_archetype_routing_table, build_enriched_evidence_registry, write_phase6_infrastructure
from src.config import DEFAULT_TARGET_YEAR, get_year_config
from src.enriched_reports import create_enriched_report, generate_enriched_report, load_context, output_filename
from src.fetch_forbes import ForbesAnnualFetcher
from src.metrics import calculate_growth_metrics
from src.pipeline import run_pipeline


ROOT = Path(__file__).resolve().parents[1]


def test_2026_year_config_paths_do_not_overlap_2025_outputs() -> None:
    cfg_2026 = get_year_config(2026)
    cfg_2025 = get_year_config(2025)

    assert cfg_2026.top100_path.as_posix().endswith("data/processed/2026/top100_2026.csv")
    assert cfg_2026.history_path.as_posix().endswith("data/processed/2026/billionaire_wealth_history_long_2026.csv")
    assert cfg_2026.metrics_path.as_posix().endswith("data/processed/2026/billionaire_growth_metrics_2026.csv")
    assert cfg_2026.citations_path.as_posix().endswith("data/processed/2026/source_citations_2026.csv")
    assert cfg_2026.people_reports_dir.as_posix().endswith("reports/people/2026")
    assert cfg_2026.top100_path != cfg_2025.top100_path
    assert cfg_2026.workbook_path.name == "Forbes_top100_2026_analysis.xlsx"


def test_current_target_year_defaults_to_2026_without_fetching() -> None:
    assert DEFAULT_TARGET_YEAR == 2026
    assert inspect.signature(run_pipeline).parameters["year"].default == DEFAULT_TARGET_YEAR
    assert ForbesAnnualFetcher().config.year == DEFAULT_TARGET_YEAR
    assert inspect.signature(load_context).parameters["year"].default == DEFAULT_TARGET_YEAR
    assert inspect.signature(create_enriched_report).parameters["year"].default == DEFAULT_TARGET_YEAR
    assert inspect.signature(generate_enriched_report).parameters["year"].default == DEFAULT_TARGET_YEAR
    assert inspect.signature(write_phase6_infrastructure).parameters["year"].default == DEFAULT_TARGET_YEAR
    assert build_archetype_routing_table()["dataset_year"].eq(DEFAULT_TARGET_YEAR).all()
    assert list(build_enriched_evidence_registry().columns)


def test_2026_growth_metrics_use_target_year_fields() -> None:
    cfg = get_year_config(2026)
    top100 = pd.DataFrame(
        [
            {
                "year": 2026,
                "rank": 1,
                "forbes_uri": "sample-founder",
                "name": "Sample Founder",
                "net_worth_usd_b": 16.0,
                "source_of_wealth": "SampleCo",
                "industry": "Technology",
                "primary_company_or_asset": "SampleCo",
                "self_made_or_inherited_if_available": "Self-made",
                "_classification_text": "founder public company technology",
            }
        ]
    )
    history = pd.DataFrame(
        [
            {"forbes_uri": "sample-founder", "name": "Sample Founder", "year": 2023, "net_worth_usd_b": 2.0},
            {"forbes_uri": "sample-founder", "name": "Sample Founder", "year": 2024, "net_worth_usd_b": 4.0},
            {"forbes_uri": "sample-founder", "name": "Sample Founder", "year": 2025, "net_worth_usd_b": 8.0},
            {"forbes_uri": "sample-founder", "name": "Sample Founder", "year": 2026, "net_worth_usd_b": 16.0},
        ]
    )

    metrics = calculate_growth_metrics(top100, history, [2023, 2024, 2025, 2026], cfg)
    row = metrics.iloc[0]

    assert "wealth_multiple_first_to_target_year" in metrics.columns
    assert "wealth_multiple_first_to_2025" not in metrics.columns
    assert row["rank"] == 1
    assert math.isclose(row["wealth_multiple_first_to_target_year"], 8.0)
    assert math.isclose(row["CAGR_nominal"], 1.0, rel_tol=1e-9)
    assert row["exponential_fit_r2"] > 0.999


def test_2026_enriched_report_filename_convention() -> None:
    cfg = get_year_config(2026)
    person = pd.Series({"rank": 1, "name": "Elon Musk", "name_slug": "elon-musk"})

    assert output_filename(person, "enriched_draft", cfg) == "001_elon-musk_business_analysis_2026_enriched_draft.docx"
    assert output_filename(person, "final_review_ready", cfg) == "001_elon-musk_business_analysis_2026_final_review_ready.docx"
    assert output_filename(person, "enriched_draft") == "001_elon-musk_business_analysis_2026_enriched_draft.docx"


def test_2026_archetype_routing_marks_2025_templates_as_references_only() -> None:
    cfg = get_year_config(2026)
    routing = build_archetype_routing_table(cfg)

    assert len(routing) >= 5
    assert routing["dataset_year"].eq(2026).all()
    assert routing["template_year"].eq(2025).all()
    assert routing["template_reference_only"].eq(True).all()


def test_2026_manual_import_templates_and_dirs_are_ready() -> None:
    cfg = get_year_config(2026)
    required_dirs = [
        cfg.raw_forbes_dir,
        cfg.interim_dir,
        cfg.processed_dir,
        cfg.annual_reports_dir,
        cfg.charts_dir,
        cfg.people_reports_dir,
    ]
    for path in required_dirs:
        assert path.exists(), path
        assert str(path).endswith("2026")

    expected_templates = {
        "manual_import_top100_2026.csv": [
            "year",
            "rank",
            "position",
            "forbes_uri",
            "name",
            "name_slug",
            "net_worth_usd_b",
            "canonical_source_url",
        ],
        "manual_import_wealth_history_2026.csv": ["forbes_uri", "name", "year", "net_worth_usd_b", "source_url"],
        "manual_import_source_citations_2026.csv": ["source_id", "field_name", "source_url", "source_mode"],
        "manual_import_person_evidence_pack_2026.csv": ["report_year", "citation_key", "claim_supported", "confidence_level"],
    }
    for file_name, expected_columns in expected_templates.items():
        path = Path("templates") / file_name
        assert path.exists(), path
        columns = list(pd.read_csv(path, nrows=0).columns)
        missing = [column for column in expected_columns if column not in columns]
        assert not missing, f"{path} missing {missing}"


def _assert_git_ignored(path: Path) -> None:
    result = subprocess.run(
        ["git", "check-ignore", "-q", str(path.as_posix())],
        cwd=ROOT,
        check=False,
    )
    assert result.returncode == 0, f"{path} is not ignored by Git"


def test_2026_generated_outputs_are_local_private_when_present() -> None:
    cfg = get_year_config(2026)
    generated_paths = [
        cfg.top100_path,
        cfg.history_path,
        cfg.metrics_path,
        cfg.citations_path,
        cfg.person_quality_path,
        cfg.workbook_path,
        cfg.charts_dir / "example.png",
        cfg.people_reports_dir / "example.docx",
        Path("data/private/2026/manual_import_top100_2026.csv"),
    ]
    for path in generated_paths:
        _assert_git_ignored(path)

    for template_name in [
        "manual_import_top100_2026.csv",
        "manual_import_wealth_history_2026.csv",
        "manual_import_source_citations_2026.csv",
        "manual_import_person_evidence_pack_2026.csv",
    ]:
        assert pd.read_csv(Path("templates") / template_name).empty

    searchable_paths = [
        path
        for path in [
            cfg.top100_path,
            cfg.history_path,
            cfg.metrics_path,
            cfg.citations_path,
            cfg.person_quality_path,
            Path("data/private/2026/manual_import_top100_2026.csv"),
            Path("data/private/2026/manual_import_wealth_history_2026.csv"),
            Path("data/private/2026/manual_import_source_citations_2026.csv"),
        ]
        if path.exists()
    ]
    text = "\n".join(path.read_text(encoding="utf-8", errors="ignore") for path in searchable_paths)
    assert "real-time" not in text.casefold()
    assert "realtime" not in text.casefold()
