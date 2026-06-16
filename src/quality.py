"""Data-quality checks and quality report generation."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .citations import HISTORY_CITED_FIELDS, TOP100_CITED_FIELDS
from .config import (
    CITATION_COLUMNS,
    DEFAULT_TARGET_YEAR,
    PEOPLE_REPORTS_DIR,
    RAW_DIR,
    REPORT_SECTIONS,
    YearConfig,
    WEALTH_ENGINE_CATEGORIES,
    get_year_config,
)


def run_quality_checks(
    top100: pd.DataFrame,
    history_long: pd.DataFrame,
    metrics: pd.DataFrame,
    source_citations: pd.DataFrame,
    fetch_failures: list[str],
    person_quality: pd.DataFrame | None = None,
    config: YearConfig | None = None,
) -> pd.DataFrame:
    """Return a structured quality-check table."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    rank_col = config.rank_col
    position_col = config.position_col
    net_worth_col = config.net_worth_col
    top100_table_name = f"top100_{config.year}"
    history_table_name = (
        "billionaire_wealth_history_long"
        if config.legacy_layout
        else f"billionaire_wealth_history_long_{config.year}"
    )
    duplicate_ranks = top100[rank_col].duplicated(keep=False)
    duplicate_rank_count = int(duplicate_ranks.sum())
    rank_note = (
        f"Duplicate rank values are Forbes ties; {position_col} remains unique."
        if duplicate_rank_count
        else "No duplicate rank values."
    )
    citation_columns_present = set(CITATION_COLUMNS).issubset(source_citations.columns)
    source_ids_unique = (
        source_citations["source_id"].astype(str).str.len().gt(0).all()
        and source_citations["source_id"].is_unique
        if "source_id" in source_citations
        else False
    )
    non_derived = (
        source_citations[source_citations["collection_method"].ne("derived_metric")]
        if "collection_method" in source_citations
        else source_citations
    )
    non_derived_urls_present = (
        non_derived["source_url"].astype(str).str.len().gt(0).all()
        if "source_url" in non_derived
        else False
    )
    citation_key_set = set()
    if {"table_name", "forbes_uri", "field_name"}.issubset(source_citations.columns):
        citation_key_set = set(
            source_citations[["table_name", "forbes_uri", "field_name"]].astype(str).agg("|".join, axis=1).tolist()
        )
    top100_coverage_missing = 0
    top100_cited_fields = TOP100_CITED_FIELDS if config.legacy_layout else [
        "year",
        "rank",
        "position",
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
    for _, row in top100.iterrows():
        uri = str(row.get("forbes_uri", ""))
        for field in top100_cited_fields:
            if f"{top100_table_name}|{uri}|{field}" not in citation_key_set:
                top100_coverage_missing += 1
    history_worth_citations = set()
    if {"table_name", "row_key", "field_name"}.issubset(source_citations.columns):
        history_worth_citations = set(
            source_citations.loc[
                source_citations["table_name"].eq(history_table_name)
                & source_citations["field_name"].eq("net_worth_usd_b"),
                "row_key",
            ].astype(str)
        )
    missing_history_worth_citations = int(
        sum(
            f"{row.forbes_uri}|{int(row.year)}" not in history_worth_citations
            for row in history_long[["forbes_uri", "year"]].itertuples(index=False)
        )
    )
    people_dir_exists = config.people_reports_dir.exists()
    raw_error_logs = sorted(path.name for path in RAW_DIR.glob("*_error.txt"))
    source_text = ""
    if not source_citations.empty:
        source_text = " ".join(
            source_citations[
                [col for col in ["source_title", "source_url", "source_type", "collection_method"] if col in source_citations]
            ]
            .fillna("")
            .astype(str)
            .agg(" ".join, axis=1)
            .tolist()
        ).lower()
    realtime_terms = ["real-time", "realtime", "real time", "rtb"]
    realtime_in_canonical = any(term in source_text for term in realtime_terms)

    checks = [
        {
            "check": "top100_exactly_100_unique_people",
            "status": "PASS" if top100["forbes_uri"].nunique() == 100 and len(top100) == 100 else "FAIL",
            "detail": f"{len(top100)} rows; {top100['forbes_uri'].nunique()} unique Forbes profile URIs.",
        },
        {
            "check": f"rank_{config.year}_duplicates",
            "status": "INFO" if duplicate_rank_count else "PASS",
            "detail": rank_note,
        },
        {
            "check": f"position_{config.year}_unique",
            "status": "PASS" if top100[position_col].is_unique else "FAIL",
            "detail": f"{top100[position_col].nunique()} unique positions.",
        },
        {
            "check": f"every_person_has_{config.year}_net_worth",
            "status": "PASS" if top100[net_worth_col].notna().all() else "FAIL",
            "detail": f"{int(top100[net_worth_col].isna().sum())} missing {config.year} net worth values.",
        },
        {
            "check": "every_top100_row_has_source_url",
            "status": "PASS" if top100["canonical_source_url"].astype(str).str.len().gt(0).all() else "FAIL",
            "detail": f"{int(top100['canonical_source_url'].astype(str).str.len().eq(0).sum())} missing canonical source URLs.",
        },
        {
            "check": f"history_has_{config.year}_record_for_each_person",
            "status": "PASS"
            if history_long[history_long["year"] == config.year]["forbes_uri"].nunique() == 100
            else "FAIL",
            "detail": (
                f"{history_long[history_long['year'] == config.year]['forbes_uri'].nunique()} "
                f"people with {config.year} history rows."
            ),
        },
        {
            "check": "growth_metrics_for_each_person",
            "status": "PASS" if metrics["forbes_uri"].nunique() == 100 and len(metrics) == 100 else "FAIL",
            "detail": f"{len(metrics)} metric rows; {metrics['forbes_uri'].nunique()} unique people.",
        },
        {
            "check": "exponential_fit_requires_enough_history",
            "status": "PASS"
            if metrics.loc[metrics["years_observed"] < 3, "exponential_fit_r2"].isna().all()
            else "FAIL",
            "detail": "R2 is only populated for rows with at least three positive annual observations.",
        },
        {
            "check": "cagr_requires_enough_history",
            "status": "PASS"
            if metrics.loc[metrics["years_observed"] < 3, "CAGR_nominal"].isna().all()
            else "FAIL",
            "detail": "CAGR is only populated for rows with at least three positive annual observations.",
        },
        {
            "check": "canonical_source_excludes_forbes_realtime",
            "status": "FAIL" if realtime_in_canonical else "PASS",
            "detail": "Canonical annual tables must not use Forbes Real-Time Billionaires source text.",
        },
        {
            "check": "wealth_engine_categories_allowed",
            "status": "PASS" if metrics["wealth_engine_category"].isin(WEALTH_ENGINE_CATEGORIES).all() else "FAIL",
            "detail": "Every wealth engine category is in the allowed AGENTS.md category list.",
        },
        {
            "check": "source_citations_present",
            "status": "PASS" if not source_citations.empty and source_citations["source_url"].notna().all() else "FAIL",
            "detail": f"{len(source_citations)} citation rows.",
        },
        {
            "check": "source_citation_schema",
            "status": "PASS" if citation_columns_present else "FAIL",
            "detail": "source_citations.csv contains the AGENTS-grade citation columns."
            if citation_columns_present
            else "source_citations.csv is missing required citation columns.",
        },
        {
            "check": "source_id_unique",
            "status": "PASS" if source_ids_unique else "FAIL",
            "detail": "Every citation has a unique nonblank source_id.",
        },
        {
            "check": "non_derived_citations_have_urls",
            "status": "PASS" if non_derived_urls_present else "FAIL",
            "detail": "Every non-derived citation has a source URL.",
        },
        {
            "check": "top100_field_citation_coverage",
            "status": "PASS" if top100_coverage_missing == 0 else "FAIL",
            "detail": f"{top100_coverage_missing} missing top100 field citation links.",
        },
        {
            "check": "history_net_worth_citation_coverage",
            "status": "PASS" if missing_history_worth_citations == 0 else "FAIL",
            "detail": f"{missing_history_worth_citations} missing history net worth citation links.",
        },
        {
            "check": "docx_report_pipeline_ready",
            "status": "PASS" if people_dir_exists else "FAIL",
            "detail": (
                f"{config.people_reports_dir.as_posix()} exists; per-person DOCX generation is command-driven, not batched."
                if people_dir_exists
                else "reports/people directory is missing."
            ),
        },
        {
            "check": "docx_report_sections_defined",
            "status": "PASS" if len(REPORT_SECTIONS) == 17 else "FAIL",
            "detail": f"{len(REPORT_SECTIONS)} report sections are configured.",
        },
        {
            "check": "forbes_fetch_failures",
            "status": "WARN" if fetch_failures else "PASS",
            "detail": "; ".join(fetch_failures[:5]) if fetch_failures else "No annual endpoint failures.",
        },
        {
            "check": "raw_fetch_error_logs_preserved",
            "status": "INFO" if raw_error_logs else "PASS",
            "detail": (
                f"{len(raw_error_logs)} preserved historical diagnostic files: {', '.join(raw_error_logs[:8])}."
                if raw_error_logs
                else "No preserved raw fetch error logs."
            ),
        },
    ]
    if person_quality is not None and not person_quality.empty:
        checks.append(
            {
                "check": "person_data_quality_scores_present",
                "status": "PASS" if person_quality["forbes_uri"].nunique() == 100 else "FAIL",
                "detail": f"{len(person_quality)} person data-quality rows.",
            }
        )
    return pd.DataFrame(checks)


def missing_field_summary(top100: pd.DataFrame, metrics: pd.DataFrame) -> pd.DataFrame:
    """List missing/uncertain values in the main processed tables."""
    rows: list[dict[str, object]] = []
    table_name = "top100_2025" if "rank_2025" in top100.columns else "top100"
    for table_name, df in [(table_name, top100), ("billionaire_growth_metrics", metrics)]:
        for col in df.columns:
            if col.startswith("_"):
                continue
            missing = int(df[col].isna().sum() + (df[col].astype(str).str.strip() == "").sum())
            if missing:
                rows.append({"table": table_name, "field": col, "missing_or_blank_rows": missing})
    return pd.DataFrame(rows)


def write_quality_report(
    path: Path,
    quality_checks: pd.DataFrame,
    missing_summary: pd.DataFrame,
    top100: pd.DataFrame,
    metrics: pd.DataFrame,
    person_quality: pd.DataFrame | None = None,
    config: YearConfig | None = None,
) -> None:
    """Write a Markdown data-quality report."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    duplicate_ranks = (
        top100.loc[top100[config.rank_col].duplicated(keep=False), [config.rank_col, config.position_col, "name"]]
        .sort_values([config.rank_col, config.position_col])
        .copy()
    )
    quality_md = quality_checks.to_markdown(index=False)
    missing_md = (
        missing_summary.to_markdown(index=False)
        if not missing_summary.empty
        else "No missing or blank values detected in required top-level fields."
    )
    duplicate_md = (
        duplicate_ranks.to_markdown(index=False)
        if not duplicate_ranks.empty
        else "No duplicate Forbes rank values detected."
    )
    low_completeness = metrics[metrics["data_completeness_score"] < 0.75][
        [config.rank_col, "name", "years_observed", "data_completeness_score"]
    ].sort_values(["data_completeness_score", config.rank_col])
    low_md = (
        low_completeness.to_markdown(index=False)
        if not low_completeness.empty
        else "No rows below 0.75 data completeness score."
    )
    person_quality_md = "Person-level quality scoring was not generated."
    if person_quality is not None and not person_quality.empty:
        person_quality_md = (
            person_quality.sort_values(["data_quality_score", config.rank_col])
            .head(15)[
                [
                    config.rank_col,
                    "name",
                    "years_observed",
                    "docx_report_ready",
                    "data_quality_score",
                    "missing_or_limited_evidence",
                ]
            ]
            .to_markdown(index=False)
        )
    text = f"""# Data Quality Report

Generated by the reproducible Forbes {config.year} Top 100 pipeline.

## Quality Checks

{quality_md}

## Missing Or Uncertain Fields

{missing_md}

## Forbes Rank Ties

Forbes annual rankings can contain ties. The project keeps the Forbes rank value and also stores a unique ordered position in the top-100 extract.

{duplicate_md}

## Low Completeness Rows

`data_completeness_score` is the share of successful annual Forbes snapshots observed from a person's first observed year through {config.year}. A late first appearance can still score highly if subsequent annual observations are complete.

{low_md}

## Person-Level Report Readiness

DOCX report readiness means the existing Forbes-derived structured fields and field-level citations are present. It does not mean that primary company filings, ownership documents, segment financials, or valuation bridge sources have already been collected.

{person_quality_md}
"""
    path.write_text(text, encoding="utf-8")
