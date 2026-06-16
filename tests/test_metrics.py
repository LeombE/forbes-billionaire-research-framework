from __future__ import annotations

import math

import pandas as pd

from src.config import get_year_config
from src.metrics import calculate_growth_metrics


def test_cagr_and_log_fit_for_smooth_compounding() -> None:
    top100 = pd.DataFrame(
        [
            {
                "forbes_uri": "sample-founder",
                "name": "Sample Founder",
                "rank_2025": 1,
                "net_worth_2025_usd_b": 8.0,
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
            {"forbes_uri": "sample-founder", "name": "Sample Founder", "year": 2022, "net_worth_usd_b": 1.0},
            {"forbes_uri": "sample-founder", "name": "Sample Founder", "year": 2023, "net_worth_usd_b": 2.0},
            {"forbes_uri": "sample-founder", "name": "Sample Founder", "year": 2024, "net_worth_usd_b": 4.0},
            {"forbes_uri": "sample-founder", "name": "Sample Founder", "year": 2025, "net_worth_usd_b": 8.0},
        ]
    )
    metrics = calculate_growth_metrics(top100, history, [2022, 2023, 2024, 2025], get_year_config(2025))
    row = metrics.iloc[0]

    assert math.isclose(row["CAGR_nominal"], 1.0, rel_tol=1e-9)
    assert row["exponential_fit_r2"] > 0.999
    assert math.isclose(row["estimated_doubling_time_years"], 1.0, rel_tol=1e-9)


def test_growth_metrics_require_three_observations() -> None:
    top100 = pd.DataFrame(
        [
            {
                "forbes_uri": "newcomer",
                "name": "New Comer",
                "rank_2025": 1,
                "net_worth_2025_usd_b": 2.0,
                "source_of_wealth": "PrivateCo",
                "industry": "Manufacturing",
                "primary_company_or_asset": "PrivateCo",
                "self_made_or_inherited_if_available": "Self-made",
                "_classification_text": "",
            }
        ]
    )
    history = pd.DataFrame(
        [
            {"forbes_uri": "newcomer", "name": "New Comer", "year": 2024, "net_worth_usd_b": 1.0},
            {"forbes_uri": "newcomer", "name": "New Comer", "year": 2025, "net_worth_usd_b": 2.0},
        ]
    )
    metrics = calculate_growth_metrics(top100, history, [2024, 2025], get_year_config(2025))

    assert pd.isna(metrics.iloc[0]["exponential_fit_r2"])
    assert pd.isna(metrics.iloc[0]["CAGR_nominal"])
    assert pd.isna(metrics.iloc[0]["max_one_year_gain_usd_b"])
