"""Phase 6A archetype routing and enriched evidence registry utilities.

This module keeps the scalable-report infrastructure reproducible without
regenerating any individual billionaire DOCX reports.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd

from .config import DEFAULT_TARGET_YEAR, INTERIM_DIR, PEOPLE_REPORTS_DIR, YearConfig, get_year_config, ensure_year_dirs
from .enriched_reports import build_evidence_items, extract_docx_validation, load_context


ROUTING_COLUMNS = [
    "wealth_engine_category",
    "archetype_name",
    "example_rank",
    "example_person",
    "required_primary_sources",
    "required_financial_fields",
    "required_ownership_fields",
    "required_charts",
    "default_confidence_level",
    "report_status_threshold",
    "template_reference_docx",
]


ENRICHED_EVIDENCE_COLUMNS = [
    "person_rank",
    "person_name",
    "source_key",
    "evidence_category",
    "source_title",
    "publisher",
    "publication_date",
    "source_url_or_locator",
    "claim_supported",
    "source_quality",
    "confidence",
    "limitation",
    "used_in_report_file",
    "accessed_at",
]


@dataclass(frozen=True)
class ReferenceReport:
    rank: int
    person_name: str
    file_name: str

    @property
    def path(self) -> Path:
        return PEOPLE_REPORTS_DIR / self.file_name

    @property
    def report_file(self) -> str:
        return f"reports/people/{self.file_name}"


REFERENCE_REPORTS = [
    ReferenceReport(1, "Elon Musk", "001_elon-musk_business_analysis_enriched_v2.docx"),
    ReferenceReport(2, "Mark Zuckerberg", "002_mark-zuckerberg_business_analysis_enriched_v3.docx"),
    ReferenceReport(3, "Jeff Bezos", "003_jeff-bezos_business_analysis_enriched_v3_clean.docx"),
    ReferenceReport(4, "Larry Ellison", "004_larry-ellison_business_analysis_enriched_v2.docx"),
    ReferenceReport(5, "Bernard Arnault & family", "005_bernard-arnault-family_business_analysis_enriched_v2.docx"),
]


ROUTING_ROWS = [
    {
        "wealth_engine_category": "Founder/operator public equity",
        "archetype_name": "High-convexity founder public equity plus private frontier optionality",
        "example_rank": 1,
        "example_person": "Elon Musk",
        "required_primary_sources": "Forbes annual-list row; public-company 10-K/annual report; proxy ownership table; official company sources; credible private-valuation sources for private assets.",
        "required_financial_fields": "Revenue; gross profit or margin; operating income; net income; operating cash flow; capex; free cash flow proxy; segment/product economics where disclosed; material private valuation marks.",
        "required_ownership_fields": "Beneficial ownership shares and percent; pledged shares; voting/control notes; option/award context; private-company ownership/valuation limits where available.",
        "required_charts": "Forbes wealth history; public-equity sensitivity; business empire map; segment economics where filings support it.",
        "default_confidence_level": "Medium-high when public-company filings support the core asset; medium/low for private optionality.",
        "report_status_threshold": "AGENTS_grade_candidate",
        "template_reference_docx": "reports/people/001_elon-musk_business_analysis_enriched_v2.docx",
    },
    {
        "wealth_engine_category": "Technology/platform monopoly/network effects",
        "archetype_name": "Public-company platform and network-effects advertising engine",
        "example_rank": 2,
        "example_person": "Mark Zuckerberg",
        "required_primary_sources": "Forbes annual-list row; 10-K/annual report; proxy ownership/voting table; investor relations release; official risk-factor disclosures.",
        "required_financial_fields": "Revenue; operating income; net income; operating cash flow; capex; free cash flow proxy; buybacks/dividends; segment revenue and operating income; user or engagement metrics where disclosed.",
        "required_ownership_fields": "Economic ownership; voting control; share classes; pledged shares; controlled-company status; trust/liquidity limitations.",
        "required_charts": "Forbes wealth history; market-cap sensitivity; platform/business empire map; segment economics or revenue mix chart.",
        "default_confidence_level": "High for filed financials and proxy control; medium for strategy synthesis and AI/option-value interpretation.",
        "report_status_threshold": "AGENTS_grade_candidate",
        "template_reference_docx": "reports/people/002_mark-zuckerberg_business_analysis_enriched_v3.docx",
    },
    {
        "wealth_engine_category": "Founder/operator public equity",
        "archetype_name": "Founder public-equity operating platform plus capital allocation and private optionality",
        "example_rank": 3,
        "example_person": "Jeff Bezos",
        "required_primary_sources": "Forbes annual-list row; 10-K/annual report; proxy ownership table; investor relations source; official private-asset/company sources when private optionality is discussed.",
        "required_financial_fields": "Revenue by segment; operating income by segment; operating cash flow; capex/PPE purchases; free cash flow proxy; buybacks/dilution; infrastructure and fulfillment cost where disclosed.",
        "required_ownership_fields": "Beneficial shares and percent; voting rights; executive chair/founder role; share-sale, trust, debt, tax, and private-asset limitations.",
        "required_charts": "Forbes wealth history; public-equity sensitivity; segment revenue/operating income; business empire map.",
        "default_confidence_level": "High for public Amazon evidence; medium-low for private assets such as Blue Origin unless primary valuation/ownership evidence exists.",
        "report_status_threshold": "AGENTS_grade_candidate",
        "template_reference_docx": "reports/people/003_jeff-bezos_business_analysis_enriched_v3_clean.docx",
    },
    {
        "wealth_engine_category": "Luxury/retail brand ownership",
        "archetype_name": "Luxury brand portfolio plus family-control wealth engine",
        "example_rank": 5,
        "example_person": "Bernard Arnault & family",
        "required_primary_sources": "Forbes annual-list row; annual report/universal registration document; ownership/voting-rights table; investor relations publications; official brand/group pages.",
        "required_financial_fields": "Revenue; profit from recurring operations; free cash flow; operating investments; net debt; segment revenue/profit; store count and regional mix where disclosed.",
        "required_ownership_fields": "Family group share capital; voting rights; controlling vehicles; intra-family allocation limits; succession/governance disclosures where available.",
        "required_charts": "Forbes wealth history; family-group sensitivity; luxury segment economics; portfolio/brand empire map.",
        "default_confidence_level": "High for group filings and family-control tables; medium/low for personal allocation, trusts, succession mechanics, and brand-level valuations.",
        "report_status_threshold": "AGENTS_grade_candidate",
        "template_reference_docx": "reports/people/005_bernard-arnault-family_business_analysis_enriched_v2.docx",
    },
    {
        "wealth_engine_category": "Founder/operator public equity",
        "archetype_name": "Enterprise software/database lock-in plus Oracle cloud infrastructure founder public-equity engine",
        "example_rank": 4,
        "example_person": "Larry Ellison",
        "required_primary_sources": "Forbes annual-list row; Oracle 10-K/annual report; Oracle proxy ownership and pledging disclosure; Oracle investor relations locator; risk-factor disclosures.",
        "required_financial_fields": "Revenue; operating margin; net income; operating cash flow; capex; free cash flow proxy; buybacks/dividends; cloud/license/support, hardware, services, and geographic revenue.",
        "required_ownership_fields": "Beneficial shares and percent; exercisable options included; pledged shares; founder/chairman/CTO role; debt/trust/liquidity limitations.",
        "required_charts": "Forbes wealth history; Oracle market-cap sensitivity; Oracle segment economics; enterprise software/cloud empire map.",
        "default_confidence_level": "High for Oracle filing/proxy facts; medium for cloud/AI valuation interpretation and personal balance-sheet bridge.",
        "report_status_threshold": "AGENTS_grade_candidate",
        "template_reference_docx": "reports/people/004_larry-ellison_business_analysis_enriched_v2.docx",
    },
]


def build_archetype_routing_table(config: YearConfig | None = None) -> pd.DataFrame:
    """Return the machine-readable archetype routing table."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    routing = pd.DataFrame(ROUTING_ROWS, columns=ROUTING_COLUMNS)
    if config.legacy_layout:
        return routing
    routing = routing.copy()
    routing.insert(0, "dataset_year", config.year)
    routing.insert(1, "route_id", [f"{config.year}-{idx + 1:03d}" for idx in range(len(routing))])
    routing["template_year"] = 2025
    routing["template_reference_only"] = True
    return routing


def _items_by_key(reference: ReferenceReport, evidence_pack_path: Path | None) -> dict[str, object]:
    context = load_context(rank=reference.rank, evidence_pack_path=evidence_pack_path)
    return {item.key: item for item in build_evidence_items(context)}


def _registry_rows_for_reference(reference: ReferenceReport, evidence_pack_path: Path | None) -> list[dict[str, object]]:
    """Build registry rows for evidence keys actually present in a reference DOCX appendix."""
    validation = extract_docx_validation(
        reference.path,
        expected_rank=reference.rank,
        expected_name=reference.person_name.split(" & ")[0],
    )
    appendix_keys = validation["appendix_evidence_keys"]
    items_by_key = _items_by_key(reference, evidence_pack_path)
    rows: list[dict[str, object]] = []
    for key in appendix_keys:
        item = items_by_key.get(key)
        if item is None:
            rows.append(
                {
                    "person_rank": reference.rank,
                    "person_name": reference.person_name,
                    "source_key": key,
                    "evidence_category": "Manual evidence pack",
                    "source_title": "Evidence key appears in DOCX appendix but is not defined in src.enriched_reports",
                    "publisher": "",
                    "publication_date": "",
                    "source_url_or_locator": reference.report_file,
                    "claim_supported": "Registry fallback row. Add a structured EvidenceItem before promoting this report further.",
                    "source_quality": "Registry fallback",
                    "confidence": "Low",
                    "limitation": "Source details must be synchronized from the DOCX appendix or evidence pack.",
                    "used_in_report_file": reference.report_file,
                    "accessed_at": "",
                }
            )
            continue
        rows.append(
            {
                "person_rank": reference.rank,
                "person_name": reference.person_name,
                "source_key": item.key,
                "evidence_category": item.category,
                "source_title": item.title,
                "publisher": item.publisher,
                "publication_date": item.date,
                "source_url_or_locator": item.locator,
                "claim_supported": item.supports,
                "source_quality": item.reliability,
                "confidence": item.confidence,
                "limitation": item.limitation,
                "used_in_report_file": reference.report_file,
                "accessed_at": item.accessed_at,
            }
        )
    return rows


def build_enriched_evidence_registry(
    *,
    references: Iterable[ReferenceReport] = REFERENCE_REPORTS,
    evidence_pack_path: Path | None = INTERIM_DIR / "ranks_3_5_evidence_pack.csv",
    config: YearConfig | None = None,
) -> pd.DataFrame:
    """Return person-specific enriched evidence rows used by reference reports."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    if not config.legacy_layout:
        return pd.DataFrame(
            columns=[
                "report_year",
                "forbes_uri",
                "person_slug",
                "person_name",
                "rank",
                "route_id",
                "template_year",
                "template_reference_only",
                "source_key",
                "claim_year",
                "source_as_of_date",
                "claim_supported",
                "source_title",
                "source_url_or_locator",
                "confidence",
                "limitation",
            ]
        )
    rows: list[dict[str, object]] = []
    for reference in references:
        rows.extend(_registry_rows_for_reference(reference, evidence_pack_path))
    registry = pd.DataFrame(rows, columns=ENRICHED_EVIDENCE_COLUMNS)
    registry = registry.drop_duplicates(["person_rank", "source_key", "used_in_report_file"]).reset_index(drop=True)
    return registry


def write_phase6_infrastructure(output_dir: Path | None = None, year: int = DEFAULT_TARGET_YEAR) -> tuple[Path, Path]:
    """Write routing and registry CSV outputs without touching DOCX reports."""
    config = get_year_config(year)
    ensure_year_dirs(config)
    output_dir = output_dir or config.interim_dir
    output_dir.mkdir(parents=True, exist_ok=True)
    suffix = "" if config.legacy_layout else f"_{config.year}"
    routing_path = output_dir / f"archetype_routing_table{suffix}.csv"
    registry_path = output_dir / f"enriched_evidence_registry{suffix}.csv"
    build_archetype_routing_table(config).to_csv(routing_path, index=False)
    build_enriched_evidence_registry(config=config).to_csv(registry_path, index=False)
    return routing_path, registry_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Write Phase 6A archetype routing and enriched evidence registry files.")
    parser.add_argument("--year", type=int, default=DEFAULT_TARGET_YEAR)
    parser.add_argument("--output-dir", type=Path, default=None)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    routing_path, registry_path = write_phase6_infrastructure(args.output_dir, year=args.year)
    routing = pd.read_csv(routing_path)
    registry = pd.read_csv(registry_path)
    print(f"Wrote archetype routing table: {routing_path} ({len(routing)} rows)")
    print(f"Wrote enriched evidence registry: {registry_path} ({len(registry)} rows)")


if __name__ == "__main__":
    main()
