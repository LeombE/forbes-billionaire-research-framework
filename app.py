"""Streamlit dashboard for annual Forbes Top 100 analysis."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from src.config import DEFAULT_TARGET_YEAR, get_year_config


BASE_DIR = Path(__file__).resolve().parent
PROCESSED_DIR = BASE_DIR / "data" / "processed"


@st.cache_data
def load_data(year: int) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    config = get_year_config(year)
    top100 = pd.read_csv(config.top100_path)
    history = pd.read_csv(config.history_path)
    metrics = pd.read_csv(config.metrics_path)
    return top100, history, metrics


def available_years() -> list[int]:
    years = [2025]
    for path in PROCESSED_DIR.glob("*/top100_*.csv"):
        try:
            years.append(int(path.parent.name))
        except ValueError:
            continue
    return sorted(set(years), reverse=True)


st.set_page_config(page_title="Forbes Top 100 Billionaires", layout="wide")
with st.sidebar:
    selected_year = st.selectbox("Annual dataset year", available_years(), index=0)

config = get_year_config(int(selected_year))
st.title(f"Forbes {config.year} Top 100 Billionaires")

try:
    top100_df, history_df, metrics_df = load_data(config.year)
except FileNotFoundError:
    st.error(f"Run `python -m src.pipeline --year {config.year} --manual-import` or `--offline` first to create processed CSV files.")
    st.stop()

rank_col = config.rank_col
net_worth_col = config.net_worth_col
data = top100_df.merge(metrics_df, on=["forbes_uri", "name", rank_col], how="left")

with st.sidebar:
    st.header("Filters")
    rank_min, rank_max = st.slider(
        f"Forbes rank {config.year}",
        int(data[rank_col].min()),
        int(data[rank_col].max()),
        (int(data[rank_col].min()), int(data[rank_col].max())),
    )
    countries = st.multiselect("Country/Territory", sorted(data["country_or_territory"].dropna().unique()))
    industries = st.multiselect("Industry", sorted(data["industry"].dropna().unique()))
    sources = st.multiselect("Source of wealth", sorted(data["source_of_wealth"].dropna().unique()))
    engines = st.multiselect("Wealth engine", sorted(data["wealth_engine_category"].dropna().unique()))
    min_cagr = st.slider("Minimum nominal CAGR", -0.25, 1.00, -0.25, 0.01)
    min_r2 = st.slider("Minimum exponential fit R^2", 0.0, 1.0, 0.0, 0.01)

filtered = data[(data[rank_col] >= rank_min) & (data[rank_col] <= rank_max)]
if countries:
    filtered = filtered[filtered["country_or_territory"].isin(countries)]
if industries:
    filtered = filtered[filtered["industry"].isin(industries)]
if sources:
    filtered = filtered[filtered["source_of_wealth"].isin(sources)]
if engines:
    filtered = filtered[filtered["wealth_engine_category"].isin(engines)]
filtered = filtered[(filtered["CAGR_nominal"].fillna(-999) >= min_cagr)]
filtered = filtered[(filtered["exponential_fit_r2"].fillna(0) >= min_r2)]

col1, col2, col3, col4 = st.columns(4)
col1.metric("People", f"{len(filtered):,}")
col2.metric("Total net worth", f"${filtered[net_worth_col].sum():,.1f}B")
col3.metric("Median CAGR", f"{filtered['CAGR_nominal'].median():.1%}" if filtered["CAGR_nominal"].notna().any() else "n/a")
col4.metric("Median R^2", f"{filtered['exponential_fit_r2'].median():.2f}" if filtered["exponential_fit_r2"].notna().any() else "n/a")

st.subheader("Filtered People")
display_cols = [
    rank_col,
    "name",
    net_worth_col,
    "country_or_territory",
    "industry",
    "source_of_wealth",
    "wealth_engine_category",
    "CAGR_nominal",
    "exponential_fit_r2",
    "data_completeness_score",
]
st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)

st.subheader("Wealth History")
selected_names = st.multiselect(
    "People to plot",
    filtered["name"].tolist(),
    default=filtered.sort_values(rank_col)["name"].head(5).tolist(),
)
selected_uris = filtered[filtered["name"].isin(selected_names)]["forbes_uri"].tolist()
plot_history = history_df[history_df["forbes_uri"].isin(selected_uris)].merge(
    top100_df[["forbes_uri", "name"]], on="forbes_uri", how="left", suffixes=("", "_canonical")
)
if not plot_history.empty:
    chart_data = plot_history.pivot_table(
        index="year",
        columns="name_canonical",
        values="net_worth_usd_b",
        aggfunc="first",
    ).sort_index()
    st.line_chart(chart_data)

st.subheader("Wealth Engine Mix")
engine_counts = filtered["wealth_engine_category"].value_counts().rename_axis("wealth_engine").reset_index(name="people")
st.bar_chart(engine_counts.set_index("wealth_engine"))
