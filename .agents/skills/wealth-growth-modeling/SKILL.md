---
name: wealth-growth-modeling
description: Use when calculating billionaire wealth history metrics: CAGR, log-linear slope, R², doubling time, drawdowns, volatility, gains/losses, and completeness scores.
---

# Wealth Growth Modeling Skill

Use this skill for quantitative wealth history analysis.

## Required formulas
- Wealth multiple = target_year_net_worth / first_net_worth.
- CAGR = (target_year_net_worth / first_net_worth) ** (1 / years_elapsed) - 1.
- Log-linear fit: ln(net_worth_usd_b) = a + b * year.
- Doubling time = ln(2) / b when b > 0.
- Annual growth = net_worth_t / net_worth_{t-1} - 1.
- Volatility = standard deviation of annual growth.
- Drawdown = net_worth / cumulative_peak - 1; largest drawdown is min drawdown.

## Rules
- Require at least 3 valid annual observations for log-linear fit unless methodology explicitly explains otherwise.
- Exclude non-positive net worth values from log fit.
- Do not call growth exponential unless R² and data coverage support it.
- Always report data completeness and limitations.
