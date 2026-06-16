"""Citation tracking and person-level data-quality scoring."""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

import numpy as np
import pandas as pd

from .config import (
    CITATION_COLUMNS,
    DEFAULT_TARGET_YEAR,
    PERSON_QUALITY_COLUMNS,
    TOP100_COLUMNS,
    YearConfig,
    annual_api_url,
    get_year_config,
)


TOP100_CITED_FIELDS = [
    "rank_2025",
    "position_2025",
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

HISTORY_CITED_FIELDS = [
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

DERIVED_METRIC_FIELDS = [
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
    "data_completeness_score",
]

TOP100_MULTIYEAR_CITED_FIELDS = [
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

DERIVED_MULTIYEAR_METRIC_FIELDS = [
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
    "data_completeness_score",
]


def _clean(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _source_id(*parts: object) -> str:
    safe = []
    for part in parts:
        text = _clean(part).lower()
        text = "".join(ch if ch.isalnum() else "-" for ch in text)
        text = "-".join(piece for piece in text.split("-") if piece)
        if text:
            safe.append(text[:40])
    return "src-" + "-".join(safe)


def _citation_row(**kwargs: Any) -> dict[str, Any]:
    row = {column: "" for column in CITATION_COLUMNS}
    row.update(kwargs)
    if not row["person_name"]:
        row["person_name"] = row.get("name", "")
    if not row["accessed_at_utc"]:
        row["accessed_at_utc"] = row.get("accessed_at", "")
    if not row["accessed_at"]:
        row["accessed_at"] = row.get("accessed_at_utc", "")
    if not row["notes"]:
        row["notes"] = row.get("evidence_note", "")
    return row


def _annual_raw_source_file(config, year: int) -> str:
    if config.legacy_layout:
        return f"data/raw/forbes_billionaires_{year}.json"
    return f"data/raw/forbes/{config.year}/forbes_billionaires_{year}.json"


def build_source_citations(
    top100: pd.DataFrame,
    history_long: pd.DataFrame,
    metrics: pd.DataFrame,
    successful_years: list[int],
    config: YearConfig | None = None,
) -> pd.DataFrame:
    """Build research-grade field/claim citation rows from project data."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    accessed_at = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    rows: list[dict[str, Any]] = []
    top100_cited_fields = TOP100_CITED_FIELDS if config.legacy_layout else TOP100_MULTIYEAR_CITED_FIELDS
    derived_metric_fields = DERIVED_METRIC_FIELDS if config.legacy_layout else DERIVED_MULTIYEAR_METRIC_FIELDS
    top100_table_name = f"top100_{config.year}"
    metrics_table_name = "billionaire_growth_metrics" if config.legacy_layout else f"billionaire_growth_metrics_{config.year}"
    history_table_name = (
        "billionaire_wealth_history_long"
        if config.legacy_layout
        else f"billionaire_wealth_history_long_{config.year}"
    )

    for year in successful_years:
        rows.append(
            _citation_row(
                source_id=_source_id("forbes", "annual", year, "dataset"),
                citation_scope="dataset_year",
                table_name="forbes_annual_raw",
                row_key=str(year),
                year=year,
                field_name="dataset_scope",
                claim_supported=f"Forbes annual billionaires list snapshot for {year}.",
                source_title=f"Forbes World's Billionaires {year} annual list JSON",
                source_url=annual_api_url(year),
                publisher="Forbes",
                author="Forbes",
                publication_date=str(year),
                accessed_at=accessed_at,
                accessed_at_utc=accessed_at,
                source_type="Official Forbes annual billionaires API",
                reliability_tier="canonical_for_forbes_list_data",
                evidence_note="Raw annual-list JSON cache used for annual ranking and net worth estimates.",
                limitations="Forbes values are estimates and annual snapshots, not audited personal balance sheets.",
                source_file=_annual_raw_source_file(config, year),
                collection_method="cached_api",
            )
        )

    for _, person in top100.iterrows():
        uri = _clean(person.get("forbes_uri"))
        name = _clean(person.get("name"))
        for field in top100_cited_fields:
            rows.append(
                _citation_row(
                    source_id=_source_id("top100", uri, field),
                    citation_scope="field",
                    person_id=uri,
                    forbes_uri=uri,
                    person_name=name,
                    name=name,
                    table_name=top100_table_name,
                    row_key=uri,
                    year=config.year,
                    field_name=field,
                    claim_supported=f"{field} for {name} in the Forbes {config.year} annual top-100 extract.",
                    source_title=f"Forbes World's Billionaires {config.year} annual list JSON",
                    source_url=_clean(person.get("canonical_source_url")),
                    publisher="Forbes",
                    author="Forbes",
                    publication_date=str(config.year),
                    accessed_at=accessed_at,
                    accessed_at_utc=accessed_at,
                    source_type="Official Forbes annual list record",
                    reliability_tier="canonical_for_forbes_list_data",
                    evidence_note=f"Field `{field}` normalized from the Forbes annual-list record for {name}.",
                    limitations="Source is a Forbes estimate/snapshot; ownership details require primary company filings.",
                    source_file=config.source_file_for_year,
                    collection_method="cached_api",
                )
            )
        rows.append(
            _citation_row(
                source_id=_source_id("profile", uri),
                citation_scope="person_profile",
                person_id=uri,
                forbes_uri=uri,
                person_name=name,
                name=name,
                table_name=top100_table_name,
                row_key=uri,
                field_name="forbes_profile_url",
                claim_supported=f"Forbes profile URL for {name}.",
                source_title=f"Forbes profile: {name}",
                source_url=_clean(person.get("forbes_profile_url")),
                publisher="Forbes",
                author="Forbes",
                publication_date="",
                accessed_at=accessed_at,
                accessed_at_utc=accessed_at,
                source_type="Forbes profile URL",
                reliability_tier="background_profile_source",
                evidence_note="Profile URL retained for audit/background review; profile prose is not copied into reports.",
                limitations="Profile content can change after the annual-list snapshot.",
                source_file="",
                collection_method="cached_api_metadata",
            )
        )

    for _, row in history_long.iterrows():
        uri = _clean(row.get("forbes_uri"))
        name = _clean(row.get("name"))
        year = int(row.get("year"))
        for field in HISTORY_CITED_FIELDS:
            rows.append(
                _citation_row(
                    source_id=_source_id("history", uri, year, field),
                    citation_scope="field",
                    person_id=uri,
                    forbes_uri=uri,
                    person_name=name,
                    name=name,
                table_name=history_table_name,
                    row_key=f"{uri}|{year}",
                    year=year,
                    field_name=field,
                    claim_supported=f"{field} for {name} in Forbes annual list year {year}.",
                    source_title=f"Forbes World's Billionaires {year} annual list JSON",
                    source_url=_clean(row.get("source_url")),
                    publisher="Forbes",
                    author="Forbes",
                    publication_date=str(year),
                    accessed_at=accessed_at,
                    accessed_at_utc=accessed_at,
                    source_type="Official Forbes annual billionaires API",
                    reliability_tier="canonical_for_forbes_list_data",
                    evidence_note=f"Field `{field}` normalized from the Forbes annual-list record for {name}, {year}.",
                    limitations="Forbes values are annual estimates; intra-year changes and methodology details are not captured.",
                    source_file=_annual_raw_source_file(config, year),
                    collection_method="cached_api",
                )
            )

    for _, metric in metrics.iterrows():
        uri = _clean(metric.get("forbes_uri"))
        name = _clean(metric.get("name"))
        for field in derived_metric_fields:
            if field not in metric or pd.isna(metric.get(field)):
                continue
            rows.append(
                _citation_row(
                    source_id=_source_id("metric", uri, field),
                    citation_scope="derived_metric",
                    person_id=uri,
                    forbes_uri=uri,
                    person_name=name,
                    name=name,
                    table_name=metrics_table_name,
                    row_key=uri,
                    year=config.year,
                    field_name=field,
                    claim_supported=f"Derived metric `{field}` for {name}.",
                    source_title="Project wealth-growth calculation from Forbes annual-list history",
                    source_url="",
                    publisher="Project pipeline",
                    author="Project pipeline",
                    publication_date=str(config.year),
                    accessed_at=accessed_at,
                    accessed_at_utc=accessed_at,
                    source_type="Derived calculation",
                    reliability_tier="derived_from_canonical_forbes_history",
                    evidence_note=f"Computed by src.metrics from annual Forbes history rows for {name}.",
                    limitations="Derived metric depends on observed annual Forbes estimates and missing-year coverage.",
                    derived_from="billionaire_wealth_history_long.net_worth_usd_b",
                    source_file=str(config.history_path.relative_to(config.history_path.parents[2])),
                    collection_method="derived_metric",
                )
            )

    citations = pd.DataFrame(rows)
    for column in CITATION_COLUMNS:
        if column not in citations:
            citations[column] = ""
    citations = citations[CITATION_COLUMNS].drop_duplicates("source_id").reset_index(drop=True)
    return citations


def build_person_data_quality_scores(
    top100: pd.DataFrame,
    history_long: pd.DataFrame,
    metrics: pd.DataFrame,
    source_citations: pd.DataFrame,
    config: YearConfig | None = None,
) -> pd.DataFrame:
    """Score report readiness and data quality at the person level."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    top100_table_name = f"top100_{config.year}"
    citation_keys = set(
        source_citations.loc[source_citations["table_name"].eq(top100_table_name), ["forbes_uri", "field_name"]]
        .astype(str)
        .agg("|".join, axis=1)
        .tolist()
    )
    history_counts = history_long.groupby("forbes_uri")["year"].nunique().to_dict()
    metric_by_uri = metrics.set_index("forbes_uri")
    rows: list[dict[str, Any]] = []

    required_top_fields = [
        field
        for field in config.top100_columns
        if field not in {config.position_col, "forbes_uri", "country_of_citizenship", "data_status", "source_quality"}
    ]
    for _, person in top100.iterrows():
        uri = _clean(person.get("forbes_uri"))
        name = _clean(person.get("name"))
        required_values_complete = all(_clean(person.get(field)) for field in required_top_fields)
        required_citations = all(f"{uri}|{field}" in citation_keys for field in required_top_fields)
        years_observed = int(history_counts.get(uri, 0))
        metric_row = metric_by_uri.loc[uri] if uri in metric_by_uri.index else pd.Series(dtype=object)
        has_growth_metrics = bool(
            years_observed >= 3
            and not pd.isna(metric_row.get("CAGR_nominal", np.nan))
            and not pd.isna(metric_row.get("exponential_fit_r2", np.nan))
        )
        completeness = float(metric_row.get("data_completeness_score", 0) or 0)
        score = (
            (0.35 if required_values_complete else 0.0)
            + min(years_observed / 10, 1.0) * 0.20
            + (0.20 if has_growth_metrics else 0.0)
            + (0.15 if required_citations else 0.0)
            + completeness * 0.10
        )
        missing_notes = []
        if not required_values_complete:
            missing_notes.append("missing required top100 fields")
        if not required_citations:
            missing_notes.append("missing required field citations")
        if not has_growth_metrics:
            missing_notes.append("insufficient history for full growth metrics")
        missing_notes.append("primary company filings/ownership data not yet collected")
        rows.append(
            {
                "year": config.year,
                "forbes_uri": uri,
                "name": name,
                config.rank_col: person.get(config.rank_col),
                "top100_required_fields_complete": required_values_complete,
                "years_observed": years_observed,
                "has_growth_metrics": has_growth_metrics,
                "has_required_citations": required_citations,
                "docx_report_ready": required_values_complete and required_citations,
                "data_quality_score": round(score, 3),
                "missing_or_limited_evidence": "; ".join(missing_notes),
            }
        )

    quality = pd.DataFrame(rows)
    for column in config.person_quality_columns:
        if column not in quality:
            quality[column] = ""
    return quality[config.person_quality_columns].sort_values(config.rank_col).reset_index(drop=True)
