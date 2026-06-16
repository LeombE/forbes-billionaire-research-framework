# MASTER CODEX PROMPT — Forbes World's Billionaires 2025 Top 100 Business Empire Research System

You are my senior Python data engineer, financial statement analyst, and business strategy research analyst.

Build and improve this repository into a reproducible research and reporting system for the Forbes World's Billionaires 2025 annual list, focused on the Top 100 people. The goal is to understand, person by person, how each fortune was built, what asset or business empire explains the ranking, what economic engine caused the net worth to compound, and what cross-industry patterns can be learned.

Do not stop after planning. First create a plan, then implement, run the pipeline, validate outputs, and report what remains manual.

## 1. Source and compliance policy
Use the official Forbes World's Billionaires 2025 annual list as the canonical source for rank, name, 2025 net worth, age, country/territory, source of wealth, industry, profile URL, and canonical source URL.

Respect robots.txt, website terms, paywalls, login gates, and rate limits. Do not bypass paywalls. If Forbes blocks automated access or data is not legally/technically available, create a manual-import workflow with templates and exact instructions. Do not fabricate missing values.

Do not mix Forbes Real-Time Billionaires data into the annual-list dataset. If real-time comparison is useful, create a separate file named `/data/processed/forbes_realtime_comparison_if_available.csv` and label it clearly.

Every factual statement in final reports must be supported by a row in `source_citations.csv` or by a generated chart/table from processed data. For public companies, prefer SEC filings, annual reports, 20-F filings, investor relations materials, exchange filings, prospectuses, and official company websites. Use reputable news/business sources only when primary sources do not answer the question.

## 2. Required repository structure
Create or repair:

```text
/data/raw/
/data/interim/
/data/processed/
/reports/
/reports/charts/
/reports/people/
/src/
/tests/
README.md
requirements.txt
methodology_notes.md
data_quality_report.md
Forbes_top100_2025_analysis.xlsx
app.py  # optional but preferred Streamlit dashboard
```

## 3. Required processed datasets
Create:

1. `/data/processed/top100_2025.csv`
2. `/data/processed/billionaire_wealth_history_long.csv`
3. `/data/processed/billionaire_growth_metrics.csv`
4. `/data/processed/source_citations.csv`

`top100_2025.csv` must contain exactly 100 unique people unless the project is in manual-import mode, in which case it must contain a clear validation failure explaining the missing fields.

## 4. Required fields for every billionaire
Collect and validate:

- rank_2025
- name
- net_worth_2025_usd_b
- age_2025
- country_or_territory
- source_of_wealth
- industry
- primary_company_or_asset
- self_made_or_inherited_if_available
- forbes_profile_url
- canonical_source_url
- notes

## 5. Wealth history and growth metrics
Build long-format annual wealth history where legally and reliably sourced. For each person, calculate growth metrics only when enough observations exist.

Required metrics:

- first_year_observed
- first_net_worth_usd_b
- years_observed
- wealth_multiple_first_to_2025
- CAGR_nominal
- log_linear_growth_slope
- exponential_fit_r2
- estimated_doubling_time_years
- peak_net_worth_usd_b
- max_one_year_gain_usd_b
- max_one_year_loss_usd_b
- volatility_of_annual_growth
- largest_drawdown_pct
- wealth_engine_category
- public_equity_dependency_flag
- key_asset_or_company
- evidence_summary
- data_completeness_score

For exponential-style growth, fit:

```text
ln(net_worth_usd_b) = a + b * year
```

Report slope, R², and doubling time = ln(2) / b when b > 0. Do not claim true exponential growth unless the data supports it. If R² is weak or observations are sparse, say so.

## 6. Wealth engine classification
Assign exactly one primary wealth engine category per person, plus confidence and evidence:

- Founder/operator public equity
- Founder/operator private company
- Early employee/executive equity
- Investor/capital allocator
- Inherited/family-controlled business
- Luxury/retail brand ownership
- Technology/platform monopoly/network effects
- Real estate/land/infrastructure
- Commodities/energy/resources
- Diversified holding company
- Other/unclear

Create a classification rubric in `methodology_notes.md`. Use transparent decision rules. If a person fits multiple categories, choose the dominant current net-worth engine and list secondary engines separately.

## 7. Individual Word reports: deep business empire analysis
Generate one Microsoft Word report per billionaire when sufficient evidence exists:

```text
/reports/people/{rank_2025:03d}_{name_slug}_business_analysis.docx
```

Each report should be as long as the available evidence justifies. For complex empires such as technology platforms, diversified holding companies, luxury groups, conglomerates, commodities, energy, real estate, or multi-entity family empires, expand the report substantially. Depth matters more than fixed length. Never pad with generic biography.

### Required sections for each person
1. Executive thesis: why this person ranks in the Top 100.
2. 2025 Forbes snapshot: rank, net worth, age, country, source of wealth, industry, profile URL.
3. Ownership and asset map: the actual equity stakes, private assets, family-controlled vehicles, trusts, foundations, voting/control rights, and major assets where sourceable.
4. Wealth history and growth metrics: annual trajectory, CAGR, exponential fit, drawdowns, volatility, and interpretation.
5. Business empire overview: what companies/assets create the wealth.
6. Segment-by-segment economics: revenue, operating income, margins, unit economics, cash flow, capex intensity, ROIC, market share, and segment growth where sourceable.
7. Financial statement linkage: explain how reported financials or valuations translate into personal net worth.
8. Industry structure: value chain, profit pools, bargaining power, regulation, supply constraints, distribution, and competitive rivalry.
9. Moat and flywheel: network effects, brand, scale, data, switching costs, IP, manufacturing capability, logistics, scarce licenses/resources, capital cost advantage, or customer lock-in.
10. First-principles founder/operator analysis: what core bottleneck was solved and why the business scaled.
11. Capital allocation: reinvestment, M&A, buybacks, leverage, holding companies, dividends, portfolio concentration, and timing.
12. Inflection-point timeline: founding, product breakthroughs, IPOs, acquisitions, market-cycle reratings, crises, and recoveries.
13. Public equity/private valuation dependency: sensitivity to stock price, valuation multiples, private rounds, FX, commodity price, or real estate cycles.
14. Country and macro context: legal system, capital market access, demographic demand, industrial policy, geopolitics, currency, tax, or regulation.
15. Risks and counter-thesis: what could shrink the fortune materially.
16. Peer comparison: compare to at least 2 relevant billionaires/companies where sourceable.
17. Strategic patterns and lessons: specific, transferable business insights.
18. Source appendix and data-quality score.

### The “why they are rich” equation
For each person, explicitly answer:

```text
Net worth ≈ ownership stake × asset value − debt/pledges/discounts + other assets
```

Then explain what caused the asset value to grow:
- Revenue growth?
- Margin expansion?
- Market share capture?
- Multiple expansion?
- Scarcity/resource inflation?
- Brand pricing power?
- Network effects?
- Regulatory monopoly/oligopoly?
- Capital allocation?
- Financial leverage?
- Inheritance/family control?

## 8. Cross-sectional insight reports
Create:

- `/reports/executive_summary.md`: 10–20 high-value findings for business readers.
- `/reports/insight_report.md`: detailed analysis of patterns by industry, country, wealth engine, public/private dependency, CAGR, exponential-fit quality, and drawdowns.
- `/reports/charts/`: charts for industry distribution, country distribution, wealth engine distribution, CAGR vs net worth, R² distribution, wealth multiples, drawdowns, and Top 20 wealth trajectory lines.

## 9. Excel workbook
Create `Forbes_top100_2025_analysis.xlsx` with sheets:

- Top100_2025
- Wealth_History_Long
- Growth_Metrics
- Industry_Summary
- Country_Summary
- Wealth_Engine_Summary
- Source_Citations
- Data_Quality
- Methodology

Use clear formatting, frozen headers, filters, numeric formats, and column widths. Validate that the workbook opens.

## 10. Optional Streamlit dashboard
Create `app.py` with filters for:
- rank
- country
- industry
- source of wealth
- wealth engine
- CAGR
- exponential_fit_r2

Dashboard should show tables, charts, and links to person reports if available.

## 11. Subagent orchestration
Spawn subagents explicitly. Use parallel agents for independent work, then synthesize. Suggested workflow:

1. Spawn `source_verification_researcher` to audit Forbes accessibility, identify canonical URLs, list manual import needs, and create `source_citations.csv` requirements.
2. Spawn `forbes_data_engineer` to inspect the existing repo, design/repair pipeline modules, tests, schemas, and CLI rerun path.
3. Spawn `wealth_history_modeler` to implement CAGR, log-linear fit, R², doubling time, volatility, and drawdown functions with tests.
4. Spawn `financial_statement_analyst` to define financial-statement extraction templates for public/private/family/holding-company cases.
5. Spawn `business_strategy_analyst` to design the first-principles report framework and wealth-engine rubric.
6. Spawn `docx_report_editor` to design Word report generation templates and citation rendering.
7. Spawn `qa_auditor` to review outputs for fabricated data, missing sources, invalid assumptions, and broken reproducibility.

Wait for all subagent results. Then implement or revise files. Do not leave the project at the plan stage.

## 12. Manual import fallback
If public Forbes annual data cannot be accessed directly, create:

```text
/data/raw/manual_import_top100_2025_template.csv
/data/raw/manual_import_wealth_history_template.csv
/data/raw/manual_import_source_citations_template.csv
```

Document the exact data I must upload, including field names, accepted formats, and examples. The pipeline must still run in manual mode and produce transparent partial outputs.

## 13. Image and chart policy
Prefer charts generated from project data. Use external images only when they add real information, such as corporate org charts, founder/product timelines, segment maps, industry value chains, or company-specific asset maps. For every external image, track source URL, license/rights status, accessed_at, caption, and alt text. Do not use decorative images. If a generated image or conceptual diagram is used, label it as AI-generated/conceptual and do not use it as evidence.

## 14. Quality checks to run
Implement and run tests for:
- exactly 100 unique people in `top100_2025.csv`
- rank uniqueness or explicit tie handling
- every person has 2025 net worth
- every row has a source URL
- valid numeric ranges for net worth, age, CAGR, R², drawdowns
- no growth metrics when there are too few observations
- no report facts without citation coverage
- all required Excel sheets exist
- all required directories and files exist

## 15. Final response after execution
At the end, report:
1. What files were created or changed.
2. How to rerun the pipeline from command line.
3. How to open the dashboard.
4. How many people have complete data.
5. Which fields still need manual sourcing.
6. Which reports were generated.
7. Any limitations caused by paywalls, robots.txt, source gaps, or uncertain values.

Begin now: inspect the repository, create a concise plan, then implement and run the pipeline. Use subagents where helpful. Do not fabricate missing Forbes or financial data.
