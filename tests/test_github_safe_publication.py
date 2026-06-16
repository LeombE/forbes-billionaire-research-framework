from __future__ import annotations

from pathlib import Path

import pandas as pd


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
        "data/raw/forbes/*/forbes_billionaires_*.json",
        "data/raw/forbes/**",
        "data/processed/*.csv",
        "data/processed/**/*.csv",
        "Forbes_top100_*_analysis.xlsx",
        "data/interim/enriched_evidence_registry*.csv",
        "data/interim/*/enriched_evidence_registry*.csv",
        "reports/people/*.docx",
        "reports/people/**/*.docx",
        "reports/charts/*.png",
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
