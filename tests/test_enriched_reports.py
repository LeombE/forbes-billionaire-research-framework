from pathlib import Path

import pandas as pd

from src.config import WEALTH_ENGINE_CATEGORIES
from src.archetype_registry import ENRICHED_EVIDENCE_COLUMNS, REFERENCE_REPORTS, ROUTING_COLUMNS
from src.enriched_reports import (
    ARCHETYPE_GUIDANCE,
    EVIDENCE_PACK_COLUMNS,
    REPORT_STATUS_VALUES,
    extract_docx_validation,
    slugify,
)


def test_person_evidence_pack_template_schema() -> None:
    template = Path("templates/person_evidence_pack_template.csv")
    assert template.exists()
    columns = list(pd.read_csv(template).columns)
    assert columns == EVIDENCE_PACK_COLUMNS


def test_all_wealth_engine_categories_have_enriched_guidance() -> None:
    missing = [category for category in WEALTH_ENGINE_CATEGORIES if category not in ARCHETYPE_GUIDANCE]
    assert missing == []
    for category in WEALTH_ENGINE_CATEGORIES:
        guidance = ARCHETYPE_GUIDANCE[category]
        assert guidance["wealth_equation"]
        assert guidance["required_evidence"]
        assert guidance["analysis_focus"]


def test_elon_enriched_v2_validation_if_generated() -> None:
    path = Path("reports/people/001_elon-musk_business_analysis_enriched_v2.docx")
    if not path.exists():
        return
    result = extract_docx_validation(path)
    assert result["valid_docx_zip"]
    assert result["missing_body_keys_from_appendix"] == []
    assert result["missing_referenced_charts"] == []
    assert not result["contains_raw_urls_in_body"]


def test_phase6_archetype_routing_outputs_if_generated() -> None:
    routing_path = Path("data/interim/archetype_routing_table.csv")
    registry_path = Path("data/interim/enriched_evidence_registry.csv")
    if not routing_path.exists() or not registry_path.exists():
        return

    routing = pd.read_csv(routing_path)
    registry = pd.read_csv(registry_path)

    assert list(routing.columns) == ROUTING_COLUMNS
    assert list(registry.columns) == ENRICHED_EVIDENCE_COLUMNS
    assert len(routing) >= 5
    assert set([1, 2, 3, 4, 5]).issubset(set(registry["person_rank"].astype(int)))
    assert registry["source_key"].astype(str).str.len().gt(0).all()
    assert registry["used_in_report_file"].astype(str).str.endswith(".docx").all()


def test_extended_docx_validation_for_larry_v2_if_generated() -> None:
    path = Path("reports/people/004_larry-ellison_business_analysis_enriched_v2.docx")
    registry_path = Path("data/interim/enriched_evidence_registry.csv")
    if not path.exists():
        return
    result = extract_docx_validation(
        path,
        expected_rank=4,
        expected_name="Larry Ellison",
        report_status="enriched_v2",
        evidence_registry_path=registry_path if registry_path.exists() else None,
        leakage_terms=["Elon", "Musk", "Tesla", "SpaceX"],
    )
    assert result["passed"]
    assert result["report_status"] in REPORT_STATUS_VALUES
    assert result["expected_rank_match"] is True
    assert result["expected_name_match"] is True
    assert result["previous_person_leakage_hits"] == {}


def test_reference_enriched_reports_validate_against_registry_if_generated() -> None:
    registry_path = Path("data/interim/enriched_evidence_registry.csv")
    if not registry_path.exists():
        return
    top100 = pd.read_csv("data/processed/top100_2025.csv")
    generated = [reference for reference in REFERENCE_REPORTS if reference.path.exists()]
    assert generated

    for reference in generated:
        person = top100[top100["rank_2025"].astype(int).eq(reference.rank)].iloc[0]
        assert slugify(str(person["name"])) in reference.path.name
        result = extract_docx_validation(
            reference.path,
            expected_rank=reference.rank,
            expected_name=str(person["name"]).split(" & ")[0],
            evidence_registry_path=registry_path,
        )
        assert result["passed"], reference.path
        assert result["body_citation_key_count"] > 0
        assert result["evidence_registry_rows"] and result["evidence_registry_rows"] > 0
        assert result["embedded_media_count"] > 0
