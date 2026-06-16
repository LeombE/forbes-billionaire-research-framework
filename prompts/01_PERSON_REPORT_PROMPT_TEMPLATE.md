# Individual Billionaire Report Prompt Template

Use this prompt when generating or revising one person’s Word report.

## Inputs
- Person row from `/data/processed/top100_2025.csv`
- Person rows from `/data/processed/billionaire_wealth_history_long.csv`
- Person row from `/data/processed/billionaire_growth_metrics.csv`
- All matching rows in `/data/processed/source_citations.csv`
- Company filings, official annual reports, investor relations materials, and reputable secondary sources already captured in citations.

## Task
Write a Microsoft Word report for `{name}` ranked `{rank_2025}` on the Forbes World's Billionaires 2025 annual list.

The report must explain why `{name}` is in the Top 100 through the actual wealth engine, not generic biography. Use financial, economic, and strategic analysis.

## Required analysis questions
1. What asset/company/equity stake explains most of the net worth?
2. How does Forbes’ 2025 net worth estimate connect to ownership stake, market value/private valuation, debt, discounts, and other assets?
3. What business model produced the asset value?
4. Which segment or product line contributes most to revenue/profit/valuation?
5. What is the key compounding mechanism: revenue growth, margins, market share, multiple expansion, capital allocation, network effects, brand pricing power, scarce resource ownership, or inheritance/family control?
6. What first-principles bottleneck did the business solve?
7. What industry structure made the business unusually profitable or defensible?
8. Which turning points caused the largest wealth gains/losses?
9. How dependent is the fortune on public equity prices or private valuation assumptions?
10. What are the major risks to the fortune?
11. What strategic lessons are transferable to entrepreneurs, operators, or investors?

## Required Word structure
Use headings and subheadings:

1. Executive thesis
2. Forbes 2025 snapshot
3. Wealth equation and asset ownership map
4. Wealth growth history and metrics
5. Business empire overview
6. Segment-by-segment business economics
7. Financial statement linkage
8. Industry structure and profit pools
9. Moat, flywheel, and first-principles analysis
10. Capital allocation and compounding
11. Inflection-point timeline
12. Public/private valuation dependency
13. Country, macro, and regulatory context
14. Risks and counter-thesis
15. Peer comparison
16. Strategic patterns and lessons
17. Source appendix and data quality

## Citation requirements
- Every paragraph containing a non-obvious factual claim must include a citation marker or source note.
- Do not cite unsupported claims.
- Do not overquote. Prefer paraphrase.
- If data is missing, write “Not available from collected sources” and list the missing source need.

## Tone
Professional business analysis. Precise, evidence-backed, specific, and useful. No motivational biography. No empty praise.
