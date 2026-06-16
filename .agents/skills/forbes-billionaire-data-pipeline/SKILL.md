---
name: forbes-billionaire-data-pipeline
description: Use when collecting, manually importing, cleaning, validating, analyzing, exporting Excel, or dashboarding annual Forbes Top 100 billionaire datasets.
---

# Forbes Billionaire Data Pipeline Skill

Use this skill for data ingestion, cleaning, validation, metrics, Excel, charts, and dashboard work.

## Workflow
1. Audit access to the official Forbes annual list for the target year.
2. If permitted and accessible, collect canonical Top 100 fields with polite rate limits.
3. If blocked or paywalled, create manual-import templates and run the pipeline in manual mode.
4. Normalize names, ranks, net worth values, countries, industries, sources, and URLs.
5. Create year-specific processed datasets and source_citations files.
6. Compute growth metrics only when there are enough observations.
7. Generate charts and Excel workbook.
8. Run tests and write data_quality_report.md.

## Hard rules
- No fabricated missing values.
- No mixing annual-list canonical data with real-time billionaire data.
- Every row must include a source URL or a manual-missing reason.
- Exactly 100 unique people are required for final pass.

## Recommended modules
- src/config.py
- src/sources/forbes_annual.py
- src/sources/manual_import.py
- src/cleaning.py
- src/metrics.py
- src/charts.py
- src/excel_export.py
- src/quality.py
- src/pipeline.py
