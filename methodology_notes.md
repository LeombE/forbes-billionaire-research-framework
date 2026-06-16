# Methodology Notes

This repository now supports annual Forbes World’s Billionaires Top 100 research across multiple years. The current target year is **2026**; the existing 2025 outputs remain preserved as the first completed historical year.

## Canonical Population

For each target year, the canonical population is the official Forbes World’s Billionaires annual list, limited to the Top 100 by ordered annual position. Forbes Real-Time Billionaires data is not used to populate annual rank or annual net worth.

## Year-Aware Outputs

Legacy 2025 files remain in their original locations. New years use explicit annual directories:

The current default target year is 2026. Legacy 2025 processing must be explicit and should be treated as historical-output maintenance, not as the default pipeline target.

- `data/raw/forbes/<year>/`
- `data/interim/<year>/`
- `data/processed/<year>/`
- `reports/annual/<year>/`
- `reports/charts/<year>/`
- `reports/people/<year>/`

## Cleaning Rules

- Convert Forbes `finalWorth` from USD millions to USD billions when using annual JSON.
- Preserve Forbes annual rank and ordered position separately.
- Keep missing values blank; do not infer age, country, source of wealth, industry, profile URL, or net worth.
- Preserve “& family” rows as Forbes publishes them.
- Manual imports must include source rows or explicit source notes.

## Growth Metrics

Metrics are calculated only where sufficient annual history exists:

- Wealth multiple = target-year net worth / first observed net worth.
- CAGR = `(target_year_net_worth / first_net_worth) ** (1 / years_elapsed) - 1`.
- Log-linear fit = `ln(net_worth_usd_b) = a + b * year`.
- Doubling time = `ln(2) / b` only when `b > 0`.
- CAGR, fit, one-year gain/loss, volatility, and drawdown require at least three valid positive annual observations.

The project reports exponential-style fit statistics, not claims of true exponential growth unless coverage and fit quality justify that wording.

## Source Hierarchy

1. Official Forbes annual list for annual rank/net worth/list fields.
2. Forbes profile URLs as background locators.
3. Primary company filings and official company documents for business/financial claims.
4. Reputable secondary sources only for background or where primary sources are unavailable.

## Legal And Publication Limits

Do not bypass paywalls, login walls, robots.txt, Cloudflare, private cookies, or rate limits. Raw Forbes files may have redistribution restrictions; the open-source repository should emphasize schemas, derived summaries, citations, and source locators rather than republishing proprietary raw content.

## Public Repository Methodology

GitHub-safe mode publishes reproducible code, methodology, blank manual-import schemas, and synthetic samples. Actual Forbes annual rows, raw API caches, processed CSVs, Excel workbooks, generated charts, and DOCX reports are local research outputs unless redistribution permission is confirmed. Synthetic samples must be clearly marked as fake and must not be used for analytical conclusions.
