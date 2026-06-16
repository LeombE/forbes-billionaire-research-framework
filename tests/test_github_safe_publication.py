from __future__ import annotations

import subprocess
from pathlib import Path

import pandas as pd
import pytest

from src.manual_templates import load_manual_citations, load_manual_inputs


ROOT = Path(__file__).resolve().parents[1]


def _gitignore_lines() -> set[str]:
    return {
        line.strip()
        for line in (ROOT / ".gitignore").read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.strip().startswith("#")
    }


def test_gitignore_excludes_forbes_derived_outputs() -> None:
    lines = _gitignore_lines()
    expected_patterns = {
        ".venv/",
        ".pytest_cache/",
        "data/raw/forbes_billionaires_*.json",
        "data/raw/forbes/2026/**",
        "data/raw/forbes/*/forbes_billionaires_*.json",
        "data/raw/forbes/**",
        "data/private/**",
        "data/processed/*.csv",
        "data/processed/2026/**",
        "data/processed/**/*.csv",
        "Forbes_top100_2026_analysis.xlsx",
        "Forbes_top100_*_analysis.xlsx",
        "data/interim/enriched_evidence_registry*.csv",
        "data/interim/*/enriched_evidence_registry*.csv",
        "reports/people/*.docx",
        "reports/people/2026/*.docx",
        "reports/people/**/*.docx",
        "reports/charts/*.png",
        "reports/charts/2026/*.png",
        "reports/charts/**/*.png",
        "reports/annual/**",
    }
    missing = expected_patterns - lines
    assert not missing
    assert not any(line == "templates/" or line.startswith("templates/manual_import") for line in lines)
    assert not any(line == "samples/" or line.startswith("samples/") for line in lines)


def test_public_manual_import_templates_remain_blank() -> None:
    for path in [
        ROOT / "templates" / "manual_import_top100_2026.csv",
        ROOT / "templates" / "manual_import_wealth_history_2026.csv",
        ROOT / "templates" / "manual_import_source_citations_2026.csv",
        ROOT / "templates" / "manual_import_person_evidence_pack_2026.csv",
    ]:
        df = pd.read_csv(path)
        assert df.empty, path


def test_synthetic_samples_match_template_headers_and_are_marked_fake() -> None:
    sample_pairs = [
        (
            ROOT / "templates" / "manual_import_top100_2026.csv",
            ROOT / "samples" / "manual_import_top100_2026_sample.csv",
        ),
        (
            ROOT / "templates" / "manual_import_wealth_history_2026.csv",
            ROOT / "samples" / "manual_import_wealth_history_2026_sample.csv",
        ),
        (
            ROOT / "templates" / "manual_import_source_citations_2026.csv",
            ROOT / "samples" / "manual_import_source_citations_2026_sample.csv",
        ),
    ]
    known_real_people = {
        "Elon Musk",
        "Mark Zuckerberg",
        "Jeff Bezos",
        "Bernard Arnault",
        "Larry Ellison",
        "Larry Page",
        "Sergey Brin",
    }
    for template_path, sample_path in sample_pairs:
        template_columns = list(pd.read_csv(template_path, nrows=0).columns)
        sample = pd.read_csv(sample_path)
        assert list(sample.columns) == template_columns
        assert len(sample) > 0
        sample_text = " ".join(sample.fillna("").astype(str).agg(" ".join, axis=1).tolist())
        assert "SYNTHETIC SAMPLE" in sample_text
        assert "example.com" in sample_text
        assert not any(real_name in sample_text for real_name in known_real_people)


def test_github_safe_policy_docs_exist_and_state_boundaries() -> None:
    files = [
        ROOT / "README.md",
        ROOT / "DATA_LICENSE_AND_SOURCE_POLICY.md",
        ROOT / "CITATION_POLICY.md",
        ROOT / "CONTRIBUTING.md",
        ROOT / "data_quality_report.md",
        ROOT / "reports" / "people" / "GITHUB_RELEASE_CHECKLIST.md",
    ]
    for path in files:
        text = path.read_text(encoding="utf-8")
        assert "synthetic" in text.lower(), path
        assert "Forbes" in text, path
    checklist = (ROOT / "reports" / "people" / "GITHUB_RELEASE_CHECKLIST.md").read_text(encoding="utf-8")
    for forbidden_output in ["raw JSON", "Processed", "Excel", "DOCX"]:
        assert forbidden_output.lower() in checklist.lower()


def test_private_manual_import_path_is_accepted(tmp_path: Path) -> None:
    top100_template = pd.read_csv(ROOT / "templates" / "manual_import_top100_2026.csv", nrows=0)
    history_template = pd.read_csv(ROOT / "templates" / "manual_import_wealth_history_2026.csv", nrows=0)
    citations_template = pd.read_csv(ROOT / "templates" / "manual_import_source_citations_2026.csv", nrows=0)

    top100_row = {column: "" for column in top100_template.columns}
    top100_row.update(
        {
            "year": 2026,
            "rank": 1,
            "position": 1,
            "forbes_uri": "synthetic-private-example",
            "name": "Synthetic Private Example",
            "name_slug": "synthetic-private-example",
            "net_worth_usd_b": 1.0,
            "source_of_wealth": "Synthetic sample",
            "industry": "Synthetic",
            "canonical_source_url": "https://example.com/private-source",
            "data_status": "synthetic_private_test",
        }
    )
    history_row = {column: "" for column in history_template.columns}
    history_row.update(
        {
            "forbes_uri": "synthetic-private-example",
            "name": "Synthetic Private Example",
            "year": 2026,
            "rank": 1,
            "position": 1,
            "net_worth_usd_b": 1.0,
            "source_url": "https://example.com/private-source",
        }
    )
    citation_row = {column: "" for column in citations_template.columns}
    citation_row.update(
        {
            "source_id": "synthetic-private-citation",
            "citation_scope": "field",
            "forbes_uri": "synthetic-private-example",
            "person_name": "Synthetic Private Example",
            "name": "Synthetic Private Example",
            "table_name": "top100_2026",
            "year": 2026,
            "field_name": "net_worth_usd_b",
            "claim_supported": "Synthetic private test claim",
            "source_url": "https://example.com/private-source",
            "source_mode": "private_manual_import",
        }
    )

    pd.DataFrame([top100_row], columns=top100_template.columns).to_csv(tmp_path / "manual_import_top100_2026.csv", index=False)
    pd.DataFrame([history_row], columns=history_template.columns).to_csv(
        tmp_path / "manual_import_wealth_history_2026.csv",
        index=False,
    )
    pd.DataFrame([citation_row], columns=citations_template.columns).to_csv(
        tmp_path / "manual_import_source_citations_2026.csv",
        index=False,
    )

    top100, history = load_manual_inputs(2026, manual_import_dir=tmp_path)
    citations = load_manual_citations(2026, manual_import_dir=tmp_path)

    assert len(top100) == 1
    assert len(history) == 1
    assert citations is not None
    assert len(citations) == 1
    assert top100.loc[0, "forbes_uri"] == "synthetic-private-example"


def test_generated_docx_reports_are_not_tracked() -> None:
    if not (ROOT / ".git").exists():
        pytest.skip("Git metadata is unavailable.")
    result = subprocess.run(
        ["git", "ls-files", "*.docx", "reports/people/*.docx", "reports/people/2026/*.docx"],
        cwd=ROOT,
        check=True,
        capture_output=True,
        text=True,
    )
    tracked_docx = [line for line in result.stdout.splitlines() if line.strip()]
    assert tracked_docx == []
