"""Aggregate summary tables for reporting and workbook sheets."""

from __future__ import annotations

import pandas as pd

from .config import DEFAULT_TARGET_YEAR, YearConfig, get_year_config


def _total_col(config: YearConfig) -> str:
    return f"total_{config.year}_net_worth_usd_b"


def _average_col(config: YearConfig) -> str:
    return f"average_{config.year}_net_worth_usd_b"


def _median_col(config: YearConfig) -> str:
    return f"median_{config.year}_net_worth_usd_b"


def build_industry_summary(top100: pd.DataFrame, metrics: pd.DataFrame, config: YearConfig | None = None) -> pd.DataFrame:
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    merged = top100.merge(metrics[["forbes_uri", "CAGR_nominal", "exponential_fit_r2"]], on="forbes_uri", how="left")
    summary = (
        merged.groupby("industry", dropna=False)
        .agg(
            people_count=("forbes_uri", "nunique"),
            **{
                _total_col(config): (config.net_worth_col, "sum"),
                _average_col(config): (config.net_worth_col, "mean"),
                _median_col(config): (config.net_worth_col, "median"),
            },
            median_CAGR_nominal=("CAGR_nominal", "median"),
            median_exponential_fit_r2=("exponential_fit_r2", "median"),
        )
        .reset_index()
        .sort_values([_total_col(config), "people_count"], ascending=False)
    )
    return summary


def build_country_summary(top100: pd.DataFrame, metrics: pd.DataFrame, config: YearConfig | None = None) -> pd.DataFrame:
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    merged = top100.merge(metrics[["forbes_uri", "CAGR_nominal"]], on="forbes_uri", how="left")
    summary = (
        merged.groupby("country_or_territory", dropna=False)
        .agg(
            people_count=("forbes_uri", "nunique"),
            **{
                _total_col(config): (config.net_worth_col, "sum"),
                _average_col(config): (config.net_worth_col, "mean"),
            },
            median_CAGR_nominal=("CAGR_nominal", "median"),
        )
        .reset_index()
        .sort_values(["people_count", _total_col(config)], ascending=False)
    )
    return summary


def build_wealth_engine_summary(top100: pd.DataFrame, metrics: pd.DataFrame, config: YearConfig | None = None) -> pd.DataFrame:
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    merged = metrics.merge(top100[["forbes_uri", config.net_worth_col]], on="forbes_uri", how="left")
    summary = (
        merged.groupby("wealth_engine_category", dropna=False)
        .agg(
            people_count=("forbes_uri", "nunique"),
            **{_total_col(config): (config.net_worth_col, "sum")},
            median_CAGR_nominal=("CAGR_nominal", "median"),
            median_exponential_fit_r2=("exponential_fit_r2", "median"),
            median_data_completeness_score=("data_completeness_score", "median"),
            public_equity_dependent_people=("public_equity_dependency_flag", "sum"),
        )
        .reset_index()
        .sort_values(["people_count", _total_col(config)], ascending=False)
    )
    return summary


def build_all_summaries(top100: pd.DataFrame, metrics: pd.DataFrame, config: YearConfig | None = None) -> dict[str, pd.DataFrame]:
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    return {
        "Industry_Summary": build_industry_summary(top100, metrics, config),
        "Country_Summary": build_country_summary(top100, metrics, config),
        "Wealth_Engine_Summary": build_wealth_engine_summary(top100, metrics, config),
    }
