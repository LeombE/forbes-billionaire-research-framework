"""Markdown reports and chart generation."""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .config import DEFAULT_TARGET_YEAR, CHARTS_DIR, REPORTS_DIR, YearConfig, get_year_config


def _fmt_b(value: float | int | None) -> str:
    if pd.isna(value):
        return "n/a"
    return f"${float(value):,.1f}B"


def _fmt_pct(value: float | int | None) -> str:
    if pd.isna(value):
        return "n/a"
    return f"{float(value) * 100:,.1f}%"


def _save_bar(df: pd.DataFrame, x: str, y: str, title: str, path: Path, *, xlabel: str = "", ylabel: str = "") -> None:
    fig_height = max(4, min(9, 0.45 * len(df) + 1.5))
    fig, ax = plt.subplots(figsize=(10, fig_height))
    ax.barh(df[x].astype(str), df[y], color="#2C7FB8")
    ax.invert_yaxis()
    ax.set_title(title)
    ax.set_xlabel(xlabel or y)
    ax.set_ylabel(ylabel or "")
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=180)
    plt.close(fig)


def create_charts(
    top100: pd.DataFrame,
    history_long: pd.DataFrame,
    metrics: pd.DataFrame,
    summaries: dict[str, pd.DataFrame],
    config: YearConfig | None = None,
) -> list[Path]:
    """Create chart images in the configured year-specific charts directory."""
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    charts_dir = config.charts_dir
    charts_dir.mkdir(parents=True, exist_ok=True)
    created: list[Path] = []
    plt.style.use("seaborn-v0_8-whitegrid")

    industry_top = summaries["Industry_Summary"].head(12)
    total_col = f"total_{config.year}_net_worth_usd_b"
    path = charts_dir / "industry_total_net_worth.png"
    _save_bar(
        industry_top.sort_values(total_col),
        "industry",
        total_col,
        f"Top 100 {config.year} Net Worth by Industry",
        path,
        xlabel="Total net worth (USD billions)",
    )
    created.append(path)

    country_top = summaries["Country_Summary"].head(12)
    path = charts_dir / "country_people_count.png"
    _save_bar(
        country_top.sort_values("people_count"),
        "country_or_territory",
        "people_count",
        f"Top 100 {config.year} People Count by Country/Territory",
        path,
        xlabel="People count",
    )
    created.append(path)

    engine_top = summaries["Wealth_Engine_Summary"]
    path = charts_dir / "wealth_engine_people_count.png"
    _save_bar(
        engine_top.sort_values("people_count"),
        "wealth_engine_category",
        "people_count",
        f"Top 100 {config.year} Count by Wealth Engine",
        path,
        xlabel="People count",
    )
    created.append(path)

    top10_uris = top100.sort_values(config.position_col).head(10)["forbes_uri"].tolist()
    top10_history = history_long[history_long["forbes_uri"].isin(top10_uris)]
    fig, ax = plt.subplots(figsize=(11, 6))
    for uri, person_history in top10_history.groupby("forbes_uri"):
        person_name = top100.loc[top100["forbes_uri"] == uri, "name"].iloc[0]
        ax.plot(person_history["year"], person_history["net_worth_usd_b"], linewidth=2, label=person_name)
    ax.set_title(f"Annual Forbes Net Worth History: {config.year} Top 10")
    ax.set_xlabel("Annual list year")
    ax.set_ylabel("Net worth (USD billions)")
    ax.legend(loc="upper left", bbox_to_anchor=(1.01, 1), fontsize=8)
    ax.grid(alpha=0.25)
    fig.tight_layout()
    path = charts_dir / "top10_wealth_history.png"
    fig.savefig(path, dpi=180)
    plt.close(fig)
    created.append(path)

    valid = metrics.dropna(subset=["CAGR_nominal", "exponential_fit_r2"])
    fig, ax = plt.subplots(figsize=(9, 6))
    if not valid.empty:
        sizes = np.clip(valid[config.multiple_col].fillna(1).to_numpy() * 8, 20, 300)
        ax.scatter(valid["CAGR_nominal"], valid["exponential_fit_r2"], s=sizes, alpha=0.65, color="#2A9D8F")
    ax.axhline(0.8, color="#E76F51", linestyle="--", linewidth=1)
    ax.set_title("CAGR vs Log-Linear Fit Quality")
    ax.set_xlabel(f"Nominal CAGR from first observed year to {config.year}")
    ax.set_ylabel("Exponential-style fit R^2")
    ax.xaxis.set_major_formatter(lambda x, _pos: f"{x:.0%}")
    ax.set_ylim(0, 1.05)
    fig.tight_layout()
    path = charts_dir / "cagr_vs_exponential_fit.png"
    fig.savefig(path, dpi=180)
    plt.close(fig)
    created.append(path)

    fig, ax = plt.subplots(figsize=(9, 5))
    cagr = metrics["CAGR_nominal"].dropna()
    if not cagr.empty:
        ax.hist(cagr, bins=18, color="#6A4C93", alpha=0.8)
    ax.set_title("Distribution of Nominal CAGR")
    ax.set_xlabel("Nominal CAGR")
    ax.set_ylabel("People count")
    ax.xaxis.set_major_formatter(lambda x, _pos: f"{x:.0%}")
    fig.tight_layout()
    path = charts_dir / "cagr_distribution.png"
    fig.savefig(path, dpi=180)
    plt.close(fig)
    created.append(path)

    return created


def write_executive_summary(
    top100: pd.DataFrame,
    metrics: pd.DataFrame,
    summaries: dict[str, pd.DataFrame],
    output_path: Path | None = None,
    config: YearConfig | None = None,
) -> None:
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    output_path = output_path or config.annual_reports_dir / "executive_summary.md"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    total_col = f"total_{config.year}_net_worth_usd_b"
    total_wealth = top100[config.net_worth_col].sum()
    top_industry = summaries["Industry_Summary"].iloc[0]
    top_country = summaries["Country_Summary"].iloc[0]
    top_engine = summaries["Wealth_Engine_Summary"].iloc[0]
    valid_cagr = metrics["CAGR_nominal"].dropna()
    high_fit = metrics[(metrics["exponential_fit_r2"] >= 0.8) & (metrics["log_linear_growth_slope"] > 0)]
    top_multiple = metrics.dropna(subset=[config.multiple_col]).sort_values(
        config.multiple_col, ascending=False
    ).head(5)

    top_multiple_md = top_multiple[
        [config.rank_col, "name", "first_year_observed", config.multiple_col, "CAGR_nominal"]
    ].to_markdown(index=False, floatfmt=".3f")

    text = f"""# Executive Summary

This project analyzes the official Forbes {config.year} annual World's Billionaires top 100 and their annual Forbes list histories from the available official annual-list JSON endpoints. It does not use Forbes Real-Time Billionaires data as the canonical annual dataset.

## Headline Findings

- The {config.year} top 100 hold {_fmt_b(total_wealth)} in combined Forbes-estimated net worth.
- The largest industry cluster is **{top_industry['industry']}**, with {int(top_industry['people_count'])} people and {_fmt_b(top_industry[total_col])}.
- The largest country/territory cluster is **{top_country['country_or_territory']}**, with {int(top_country['people_count'])} people.
- The most common wealth-engine category is **{top_engine['wealth_engine_category']}**, with {int(top_engine['people_count'])} people.
- Median nominal CAGR among rows with enough history is {_fmt_pct(valid_cagr.median() if not valid_cagr.empty else np.nan)}.
- {len(high_fit)} people have a positive log-linear wealth slope with R^2 of at least 0.80. That is evidence of exponential-style compounding in the observed Forbes estimates, not proof of true exponential wealth growth.

## Largest Observed Multiples

{top_multiple_md}

## Practical Interpretation

The most repeatable pattern is not simply "high income." It is control of scarce equity-like assets: founder stakes, family-controlled operating companies, concentrated public-company shares, capital-allocation vehicles, and real assets that scale through market cycles. The top 100 mostly reached large capital through ownership concentration, not salary accumulation.

## Source Notes

- Canonical ranking, net worth, country, age, industry, and source-of-wealth fields come from Forbes annual-list records or manual annual-list imports for {config.year}.
- Growth metrics are derived by `src.metrics` from `{config.history_path.as_posix()}`.
- Field-level citation rows are available in `{config.citations_path.as_posix()}`.
- Company financial statements, ownership percentages, segment economics, and valuation bridges have not yet been collected; those claims are intentionally excluded or labeled as future source needs.
"""
    output_path.write_text(text, encoding="utf-8")


def write_insight_report(
    top100: pd.DataFrame,
    history_long: pd.DataFrame,
    metrics: pd.DataFrame,
    summaries: dict[str, pd.DataFrame],
    output_path: Path | None = None,
    config: YearConfig | None = None,
) -> None:
    config = config or get_year_config(DEFAULT_TARGET_YEAR)
    output_path = output_path or config.annual_reports_dir / "insight_report.md"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    industry_md = summaries["Industry_Summary"].head(10).to_markdown(index=False, floatfmt=".3f")
    country_md = summaries["Country_Summary"].head(10).to_markdown(index=False, floatfmt=".3f")
    engine_md = summaries["Wealth_Engine_Summary"].to_markdown(index=False, floatfmt=".3f")

    drawdowns = metrics.dropna(subset=["largest_drawdown_pct"]).sort_values("largest_drawdown_pct").head(10)
    drawdowns_md = drawdowns[
        ["rank_2025", "name", "largest_drawdown_pct", "volatility_of_annual_growth", "wealth_engine_category"]
        if config.legacy_layout
        else ["rank", "name", "largest_drawdown_pct", "volatility_of_annual_growth", "wealth_engine_category"]
    ].to_markdown(index=False, floatfmt=".3f")

    strongest_fit = metrics.dropna(subset=["exponential_fit_r2", "log_linear_growth_slope"]).sort_values(
        ["exponential_fit_r2", "log_linear_growth_slope"], ascending=False
    ).head(10)
    strongest_fit_md = strongest_fit[
        [
            config.rank_col,
            "name",
            "log_linear_growth_slope",
            "exponential_fit_r2",
            "estimated_doubling_time_years",
            "wealth_engine_category",
        ]
    ].to_markdown(index=False, floatfmt=".3f")

    top_cagr = metrics.dropna(subset=["CAGR_nominal"]).sort_values("CAGR_nominal", ascending=False).head(10)
    top_cagr_md = top_cagr[
        [config.rank_col, "name", "first_year_observed", "CAGR_nominal", config.multiple_col]
    ].to_markdown(index=False, floatfmt=".3f")

    public_share = metrics["public_equity_dependency_flag"].mean()
    median_history_years = metrics["years_observed"].median()
    text = f"""# Insight Report

## Scope

The canonical population is the Forbes World's Billionaires {config.year} annual list, ordered by Forbes position and limited to the top 100 rows. Historical values come from annual Forbes list snapshots or manual annual-list imports, not real-time rankings.

## Strategic Patterns

1. **Ownership beats income.** Most fortunes are tied to equity stakes, family-controlled enterprises, investment vehicles, or scarce real assets. Salary-like paths are largely absent.
2. **Concentration matters.** The largest fortunes usually reflect a concentrated position in one compounding asset, then optional diversification after the core asset reaches scale.
3. **Public markets are an amplifier.** About {_fmt_pct(public_share)} of the top 100 have a rule-based public-equity dependency flag. This makes annual Forbes net worth estimates sensitive to market prices and explains large one-year swings.
4. **Exponential-style growth is uneven.** A high R^2 in the log-linear model indicates that the observed annual Forbes estimates resemble exponential compounding, but step changes, IPOs, stock re-ratings, dilution, donations, divorces, and currency effects often break a smooth curve.
5. **Family control remains a major route.** Inherited or family-controlled rows are common enough that "wealth engine" should not be read only as founder entrepreneurship.

Median observed annual history length is {median_history_years:.0f} Forbes annual rows.

## Industry Summary

{industry_md}

## Country/Territory Summary

{country_md}

## Wealth Engine Summary

{engine_md}

## Highest Nominal CAGR

{top_cagr_md}

## Strongest Exponential-Style Fits

These are the strongest log-linear fits in the annual Forbes estimates. They are candidates for exponential-style compounding, not definitive proof of true exponential wealth creation.

{strongest_fit_md}

## Largest Drawdowns

Drawdown is measured from the prior observed peak in annual Forbes net worth estimates.

{drawdowns_md}

## Limitations

- Forbes annual net worth values are estimates and snapshots, not audited personal balance sheets.
- The Forbes {config.year} list can include people labelled "& family"; this project does not split family fortunes into individuals.
- Historical matching uses Forbes profile URI. If Forbes changed a person's URI in older data, that history may be incomplete.
- Classification is rule-based and transparent, but some people fit multiple categories. The assigned category is the dominant rule match, not a claim that no other engine matters.
- Forbes real-time data is intentionally excluded.

## Source Notes

- Aggregate tables in this report are derived from `{config.top100_path.as_posix()}`, `{config.history_path.as_posix()}`, and `{config.metrics_path.as_posix()}`.
- Source IDs, field names, URLs, publishers, access timestamps, evidence notes, and limitations are tracked in `{config.citations_path.as_posix()}`.
- Forbes annual-list records are canonical for Forbes list fields; primary company filings are still needed before final per-person financial statement claims.
"""
    output_path.write_text(text, encoding="utf-8")
