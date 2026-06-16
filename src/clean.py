"""Normalize Forbes annual-list JSON into analysis tables."""

from __future__ import annotations

import html
import re
import unicodedata
from typing import Any

import numpy as np
import pandas as pd

from .config import (
    CANONICAL_YEAR,
    DEFAULT_TARGET_YEAR,
    HISTORY_COLUMNS,
    TOP100_COLUMNS,
    YearConfig,
    annual_api_url,
    get_year_config,
    profile_url,
)
from .fetch_forbes import extract_records


def _clean_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    return " ".join(html.unescape(str(value)).replace("\r", " ").replace("\n", " ").split())


def _slugify(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", normalized).strip("-").lower()
    return slug or "person"


def _first_industry(record: dict[str, Any]) -> str:
    industries = record.get("industries")
    if isinstance(industries, list) and industries:
        return _clean_text(industries[0])
    return _clean_text(record.get("category"))


def _primary_asset(record: dict[str, Any]) -> str:
    employment = record.get("employment")
    if isinstance(employment, dict):
        name = _clean_text(employment.get("name"))
        if name:
            return name
    organization = _clean_text(record.get("organization"))
    if organization:
        return organization
    return _clean_text(record.get("source"))


def _self_made_label(value: Any) -> str:
    if value is True:
        return "Self-made"
    if value is False:
        return "Inherited/family/other"
    return "Unknown"


def _worth_usd_b(record: dict[str, Any]) -> float:
    worth = record.get("finalWorth")
    if worth is None:
        return np.nan
    return round(float(worth) / 1000.0, 3)


def _sort_key(record: dict[str, Any]) -> tuple[int, int, str]:
    position = record.get("position")
    rank = record.get("rank")
    return (
        int(position) if position is not None else 999_999,
        int(rank) if rank is not None else 999_999,
        _clean_text(record.get("personName")),
    )


def build_top100(payload: dict[str, Any], config: YearConfig | None = None) -> pd.DataFrame:
    """Build the canonical top-100 table for one annual Forbes payload."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    records = sorted(extract_records(payload), key=_sort_key)[:100]
    rows: list[dict[str, Any]] = []
    source_url = annual_api_url(config.year, limit=100)
    for record in records:
        uri = _clean_text(record.get("uri") or record.get("person", {}).get("uri"))
        position = record.get("position")
        rank = record.get("rank")
        rank_note = "Forbes rank tie" if position != rank else ""
        name = _clean_text(record.get("personName") or record.get("person", {}).get("name"))
        common = {
            "forbes_uri": uri,
            "name": name,
            "country_or_territory": _clean_text(record.get("country")),
            "country_of_citizenship": _clean_text(record.get("countryOfCitizenship")),
            "source_of_wealth": _clean_text(record.get("source")),
            "industry": _first_industry(record),
            "primary_company_or_asset": _primary_asset(record),
            "self_made_or_inherited_if_available": _self_made_label(record.get("selfMade")),
            "forbes_profile_url": profile_url(uri),
            "canonical_source_url": source_url,
            "notes": _clean_text(
                f"Forbes {config.year} annual-list record. Position {position}; rank {rank}. {rank_note}"
            ),
            "_classification_text": _clean_text(
                " ".join(
                    [
                        str(record.get("source", "")),
                        str(record.get("category", "")),
                        str(record.get("organization", "")),
                        " ".join(record.get("bios", []) if isinstance(record.get("bios"), list) else []),
                        str(record.get("bio", "")),
                    ]
                )
            ),
        }
        if config.legacy_layout:
            row = {
                "rank_2025": rank,
                "position_2025": position,
                "net_worth_2025_usd_b": _worth_usd_b(record),
                "age_2025": record.get("age"),
                **common,
            }
        else:
            row = {
                "year": config.year,
                "rank": rank,
                "position": position,
                "name_slug": _slugify(name),
                "net_worth_usd_b": _worth_usd_b(record),
                "age": record.get("age"),
                "canonical_annual_list_source_url": source_url,
                "data_status": "collected_from_forbes_annual",
                "source_quality": "canonical_forbes_annual_list",
                **common,
            }
        rows.append(row)

    df = pd.DataFrame(rows)
    for col in config.top100_columns:
        if col not in df:
            df[col] = np.nan
    return df[config.top100_columns + ["_classification_text"]]


def build_top100_2025(payload_2025: dict[str, Any]) -> pd.DataFrame:
    """Build the canonical 2025 top-100 table from the annual Forbes payload."""
    return build_top100(payload_2025, get_year_config(2025))


def build_wealth_history_long(
    year_payloads: dict[int, dict[str, Any]],
    top100: pd.DataFrame,
    config: YearConfig | None = None,
) -> pd.DataFrame:
    """Build a long annual wealth-history table for the target-year top 100."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    top_uris = set(top100["forbes_uri"].dropna().astype(str))
    display_names = top100.set_index("forbes_uri")["name"].to_dict()
    rows: list[dict[str, Any]] = []

    for year, payload in sorted(year_payloads.items()):
        source_url = annual_api_url(int(year))
        for record in extract_records(payload):
            uri = _clean_text(record.get("uri") or record.get("person", {}).get("uri"))
            if uri not in top_uris:
                continue
            worth_b = _worth_usd_b(record)
            if pd.isna(worth_b) or worth_b <= 0:
                continue
            rows.append(
                {
                    "forbes_uri": uri,
                    "name": display_names.get(
                        uri,
                        _clean_text(record.get("personName") or record.get("person", {}).get("name")),
                    ),
                    "year": int(year),
                    "rank": record.get("rank"),
                    "position": record.get("position"),
                    "net_worth_usd_b": worth_b,
                    "age": record.get("age"),
                    "country_or_territory": _clean_text(record.get("country")),
                    "source_of_wealth": _clean_text(record.get("source")),
                    "industry": _first_industry(record),
                    "primary_company_or_asset": _primary_asset(record),
                    "forbes_profile_url": profile_url(uri),
                    "source_url": source_url,
                }
            )

    history = pd.DataFrame(rows)
    for col in HISTORY_COLUMNS:
        if col not in history:
            history[col] = np.nan
    if history.empty:
        return history[HISTORY_COLUMNS]

    positions = top100.set_index("forbes_uri")[config.position_col].to_dict()
    history["_target_position"] = history["forbes_uri"].map(positions)
    history = history.sort_values(["_target_position", "year", "name"]).drop(columns=["_target_position"])
    return history[HISTORY_COLUMNS].reset_index(drop=True)
