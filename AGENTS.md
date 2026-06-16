# AGENTS.md - Multi-Year Forbes Top 100 Billionaires Research Project

## Mission

This is a reproducible Python research project for annual Forbes World’s Billionaires Top 100 analysis. The current target year is 2026, while the completed 2025 work remains preserved as historical output and archetype reference material.

The project must produce structured data, quality checks, Excel workbooks, charts, and long-form business-analysis reports that explain wealth engines through ownership, assets, business economics, industry structure, capital markets, and strategic decisions.

## Source Policy

- Use the official Forbes World’s Billionaires annual list as the canonical ranking source for each target year.
- Do not mix Forbes Real-Time Billionaires data into the canonical annual dataset.
- If real-time data is useful, store it only in a separate comparison file and label it clearly.
- Respect robots.txt, paywalls, terms, login walls, rate limits, and access controls.
- If official annual access is blocked or unclear, use manual-import templates and document exactly which fields the user must provide.
- Do not fabricate missing values.
- Every material factual claim must trace to a source URL, source file, or source note.
- Prefer primary company sources for filings, ownership, financial statements, valuation bridges, and risk factors.

## Required Multi-Year Structure

- `config/project.yml`
- `data/raw/forbes/<year>/`
- `data/interim/<year>/`
- `data/processed/<year>/`
- `reports/annual/<year>/`
- `reports/charts/<year>/`
- `reports/people/<year>/`
- `templates/manual_import_*_<year>.csv`

Legacy 2025 flat paths are preserved for backward compatibility.

## Target-Year Processed Outputs

For 2026, expect:

- `data/processed/2026/top100_2026.csv`
- `data/processed/2026/billionaire_wealth_history_long_2026.csv`
- `data/processed/2026/billionaire_growth_metrics_2026.csv`
- `data/processed/2026/source_citations_2026.csv`
- `data/interim/2026/enriched_evidence_registry_2026.csv`
- `data/interim/2026/archetype_routing_table_2026.csv`
- `Forbes_top100_2026_analysis.xlsx`

## Growth Metrics

Calculate only when enough history exists:

- first year observed
- first net worth
- years observed
- wealth multiple first-to-target-year
- nominal CAGR
- log-linear slope
- R^2
- doubling time when slope > 0
- peak net worth
- max one-year gain/loss
- annual growth volatility
- largest drawdown
- data completeness score

Use `ln(net_worth_usd_b) = a + b * year`. Do not claim true exponential growth unless data coverage and fit quality support it.

## Report Rules

Per-person reports belong under `reports/people/<year>/`.

Filename format:

```text
001_<name_slug>_business_analysis_<year>_enriched_draft.docx
```

Upgrades use `_enriched_v2.docx` or `_final_review_ready.docx`. Do not overwrite existing versions without explicit approval.

Every enriched report should include:

- Title page
- How to read this report
- Report map
- Executive thesis
- Wealth equation and asset map
- Wealth history and exponential-style fit
- Business empire timeline
- Archetype-specific lens
- Core asset/company/portfolio analysis
- Financial statement or valuation linkage
- Industry and competitive position
- Moat/flywheel analysis
- Capital allocation pattern
- Country, macro, regulatory, and geopolitical context
- Risks and counter-thesis
- Comparable billionaire patterns
- Transferable lessons
- Evidence appendix, evidence gaps, confidence levels, claims-not-final, and limitations

## Archetype Policy

The 2025 rank 1-5 enriched reports are template references only:

- Elon Musk
- Mark Zuckerberg
- Jeff Bezos
- Bernard Arnault & family
- Larry Ellison

Do not copy their person-specific text into 2026 reports. If a new 2026 person does not fit the existing archetypes, stop and create a new archetype template before batch generation.

## Subagent Protocol

For large tasks, use these roles deliberately:

- `forbes_data_engineer`
- `source_verification_researcher`
- `wealth_history_modeler`
- `financial_statement_analyst`
- `business_strategy_analyst`
- `docx_report_editor`
- `qa_auditor`

Subagents must not fabricate missing data. They must return citations or mark evidence unavailable.

## Done Means

- Year-aware pipeline commands run.
- Tests pass.
- 2025 outputs are not overwritten by 2026 work.
- Manual-import templates exist when canonical data is unavailable.
- Documentation explains source policy, methodology, limitations, and annual update workflow.
- Final response reports files changed, commands run, tests, manual import status, and next command for rank 1.
