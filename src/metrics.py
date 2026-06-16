"""Growth metric calculations for billionaire wealth histories."""

from __future__ import annotations

import math

import numpy as np
import pandas as pd

from .classify import classify_wealth_engine
from .config import DEFAULT_TARGET_YEAR, GROWTH_COLUMNS, YearConfig, get_year_config


def _safe_float(value: object) -> float:
    try:
        if pd.isna(value):
            return np.nan
        return float(value)
    except (TypeError, ValueError):
        return np.nan


def _log_linear_fit(history: pd.DataFrame) -> tuple[float, float, float]:
    """Fit ln(net worth) = a + b * year. Return slope, r2, doubling time."""
    valid = history.dropna(subset=["year", "net_worth_usd_b"]).copy()
    valid = valid[valid["net_worth_usd_b"] > 0]
    if len(valid) < 3:
        return np.nan, np.nan, np.nan

    x = valid["year"].astype(float).to_numpy()
    y = np.log(valid["net_worth_usd_b"].astype(float).to_numpy())
    slope, intercept = np.polyfit(x, y, 1)
    predicted = intercept + slope * x
    ss_res = float(np.sum((y - predicted) ** 2))
    ss_tot = float(np.sum((y - np.mean(y)) ** 2))
    r2 = np.nan if ss_tot == 0 else 1.0 - ss_res / ss_tot
    doubling_time = math.log(2) / slope if slope > 0 else np.nan
    return float(slope), float(r2), float(doubling_time)


def _largest_drawdown_pct(values: pd.Series) -> float:
    if values.empty:
        return np.nan
    running_peak = values.cummax()
    drawdowns = (values - running_peak) / running_peak
    return float(drawdowns.min())


def _annual_change_metrics(history: pd.DataFrame) -> dict[str, float]:
    ordered = history.sort_values("year").copy()
    ordered["year_delta"] = ordered["year"].diff()
    ordered["worth_delta"] = ordered["net_worth_usd_b"].diff()
    ordered["growth_pct"] = ordered["net_worth_usd_b"].pct_change()
    adjacent = ordered[ordered["year_delta"] == 1]
    if adjacent.empty:
        return {
            "max_one_year_gain_usd_b": np.nan,
            "max_one_year_loss_usd_b": np.nan,
            "volatility_of_annual_growth": np.nan,
        }

    gains = adjacent["worth_delta"].dropna()
    growth = adjacent["growth_pct"].replace([np.inf, -np.inf], np.nan).dropna()
    max_gain = float(gains.max()) if not gains.empty else np.nan
    max_loss = float(gains.min()) if not gains.empty else np.nan
    volatility = float(growth.std(ddof=1)) if len(growth) >= 2 else np.nan
    return {
        "max_one_year_gain_usd_b": max_gain,
        "max_one_year_loss_usd_b": max_loss,
        "volatility_of_annual_growth": volatility,
    }


def _completeness_score(history: pd.DataFrame, successful_years: list[int], target_year: int) -> float:
    if history.empty or target_year not in set(history["year"].astype(int)):
        return 0.0
    first_year = int(history["year"].min())
    expected_years = [year for year in successful_years if first_year <= year <= target_year]
    if not expected_years:
        return 0.0
    observed = history["year"].nunique()
    return round(float(observed / len(expected_years)), 3)


def calculate_growth_metrics(
    top100: pd.DataFrame,
    history_long: pd.DataFrame,
    successful_years: list[int],
    config: YearConfig | None = None,
) -> pd.DataFrame:
    """Calculate required growth metrics for every target-year top-100 record."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    rows: list[dict[str, object]] = []
    history_by_uri = {uri: df.copy() for uri, df in history_long.groupby("forbes_uri")}

    for _, person in top100.iterrows():
        uri = str(person.get("forbes_uri", ""))
        person_history = history_by_uri.get(uri, pd.DataFrame(columns=history_long.columns))
        person_history = person_history.dropna(subset=["year", "net_worth_usd_b"]).copy()
        person_history = person_history[person_history["net_worth_usd_b"] > 0].sort_values("year")

        net_target = _safe_float(person.get(config.net_worth_col))
        first_year = np.nan
        first_net = np.nan
        span = np.nan
        multiple = np.nan
        cagr = np.nan
        peak = np.nan
        drawdown = np.nan
        annual_metrics = {
            "max_one_year_gain_usd_b": np.nan,
            "max_one_year_loss_usd_b": np.nan,
            "volatility_of_annual_growth": np.nan,
        }
        slope, r2, doubling_time = np.nan, np.nan, np.nan

        if not person_history.empty:
            first = person_history.iloc[0]
            first_year = int(first["year"])
            first_net = float(first["net_worth_usd_b"])
            span = int(config.year - first_year)
            peak = float(person_history["net_worth_usd_b"].max())
            years_observed = int(person_history["year"].nunique())

            if years_observed >= 3:
                drawdown = _largest_drawdown_pct(person_history["net_worth_usd_b"].astype(float))
                annual_metrics = _annual_change_metrics(person_history)
                slope, r2, doubling_time = _log_linear_fit(person_history)

                if span > 0 and first_net > 0 and net_target > 0:
                    multiple = float(net_target / first_net)
                    cagr = float((net_target / first_net) ** (1 / span) - 1)
                elif span == 0:
                    multiple = 1.0

        engine = classify_wealth_engine(person)
        row = {
            "forbes_uri": uri,
            "name": person.get("name", ""),
            "first_year_observed": first_year,
            "first_net_worth_usd_b": first_net,
            "years_observed": int(person_history["year"].nunique()) if not person_history.empty else 0,
            "observation_span_years": span,
            "CAGR_nominal": cagr,
            "log_linear_growth_slope": slope,
            "exponential_fit_r2": r2,
            "estimated_doubling_time_years": doubling_time,
            "peak_net_worth_usd_b": peak,
            "largest_drawdown_pct": drawdown,
            "wealth_engine_category": engine.category,
            "secondary_wealth_engines": engine.secondary_wealth_engines,
            "classification_confidence": engine.classification_confidence,
            "public_equity_dependency_flag": engine.public_equity_dependency_flag,
            "key_asset_or_company": engine.key_asset_or_company,
            "evidence_summary": engine.evidence_summary,
            "source_ids": f"src-top100-{uri}-source-of-wealth; src-top100-{uri}-industry; src-top100-{uri}-primary-company-or-asset",
            "data_completeness_score": _completeness_score(person_history, successful_years, config.year),
        }
        if config.legacy_layout:
            row["rank_2025"] = person.get(config.rank_col, np.nan)
            row["wealth_multiple_first_to_2025"] = multiple
        else:
            row["year"] = config.year
            row["rank"] = person.get(config.rank_col, np.nan)
            row["wealth_multiple_first_to_target_year"] = multiple
        row.update(annual_metrics)
        rows.append(row)

    metrics = pd.DataFrame(rows)
    for col in config.growth_columns:
        if col not in metrics:
            metrics[col] = np.nan
    return metrics[config.growth_columns].sort_values([config.rank_col, "name"]).reset_index(drop=True)
