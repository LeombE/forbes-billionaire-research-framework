# Forbes World’s Billionaires Top 100 Research Project

Open-source-ready Python research project for annual Forbes World’s Billionaires Top 100 analysis. The current target year is **2026**; the existing 2025 processed data, workbook, charts, and rank 1-5 enriched reports are preserved as historical outputs and archetype references.

## Source Policy

- Canonical ranking source: official Forbes World’s Billionaires annual list for the target year.
- Canonical annual fields: rank, ordered position, net worth, country/territory, source of wealth, industry, Forbes profile URL, and annual-list source URL.
- Forbes Real-Time Billionaires data is not canonical. If used, it must stay in a separate comparison file such as `data/processed/2026/forbes_realtime_comparison_june_2026.csv`.
- Do not bypass paywalls, login walls, robots.txt, rate limits, or access controls.
- If official annual access is blocked or not permitted, use manual-import templates and leave unavailable fields blank. Do not fabricate missing values.

## GitHub-Safe Publication Mode

This repository is configured to publish code, documentation, blank manual-import templates, and synthetic samples only. Do not commit Forbes raw JSON, populated annual CSVs, processed Forbes-derived outputs, Excel workbooks, generated charts, enriched evidence registries, or billionaire DOCX reports. Local outputs may remain in the workspace, but `.gitignore` keeps them out of the public repo.

Synthetic examples live under `samples/` and use fake people, fake companies, fake URLs, and source notes. They demonstrate schema shape only and are not Forbes data.

## License

This repository is released under the MIT License for the project code, documentation, blank templates, and synthetic sample files only.

Forbes data, third-party source data, generated Forbes-derived datasets, generated Excel workbooks, generated charts, generated DOCX reports, and user-provided annual datasets are not included in this license grant. Bring your own legally usable source data before running the pipeline.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Annual Workflow

The current default target year in Python entrypoints is 2026. Use `--year 2025` only for explicit legacy maintenance; preserved 2025 flat outputs are not regenerated during 2026 migration work.

Create target-year folders and manual-import templates:

```powershell
.\.venv\Scripts\python.exe -m src.pipeline --year 2026 --init-year
```

Run the pipeline from cached annual JSON:

```powershell
.\.venv\Scripts\python.exe -m src.pipeline --year 2026 --offline
```

Run from filled manual-import files:

```powershell
.\.venv\Scripts\python.exe -m src.pipeline --year 2026 --manual-import
```

Refresh official annual data only after confirming access is permitted:

```powershell
.\.venv\Scripts\python.exe -m src.pipeline --year 2026 --force-fetch
```

Validate project tests:

```powershell
.\.venv\Scripts\python.exe -m pytest
```

## 2026 Manual Import Files

If Forbes annual data cannot be accessed programmatically, fill these files:

- `templates/manual_import_top100_2026.csv`
- `templates/manual_import_wealth_history_2026.csv`
- `templates/manual_import_source_citations_2026.csv`
- `templates/manual_import_person_evidence_pack_2026.csv`

Every manually entered annual field must have a source row or source note. Use official Forbes annual list fields as canonical; do not use Forbes Real-Time values to populate annual rank or annual net worth.

For public GitHub releases, keep these manual-import files blank. Populate them only in a private local workspace after confirming source permissions.

## Expected 2026 Outputs

- `data/processed/2026/top100_2026.csv`
- `data/processed/2026/billionaire_wealth_history_long_2026.csv`
- `data/processed/2026/billionaire_growth_metrics_2026.csv`
- `data/processed/2026/source_citations_2026.csv`
- `data/interim/2026/enriched_evidence_registry_2026.csv`
- `data/interim/2026/archetype_routing_table_2026.csv`
- `reports/annual/2026/`
- `reports/charts/2026/`
- `reports/people/2026/`
- `Forbes_top100_2026_analysis.xlsx`

The 2025 legacy outputs remain in their existing flat paths for backward compatibility.

## Generate One Enriched DOCX

Do not batch-generate all 100 reports. Generate one report, validate it, then proceed in small batches only when evidence coverage is adequate.

```powershell
.\.venv\Scripts\python.exe -m src.enriched_reports --year 2026 --rank 1 --variant enriched_draft
```

Expected filename:

```text
reports/people/2026/001_<name_slug>_business_analysis_2026_enriched_draft.docx
```

Upgraded versions must use new filenames such as `_enriched_v2.docx` or `_final_review_ready.docx`; do not overwrite previous versions unless explicitly approved.

## Archetypes

The 2025 rank 1-5 enriched reports are reusable archetype references only:

- Elon Musk: founder public equity plus private frontier optionality
- Mark Zuckerberg: platform/network-effects advertising engine
- Jeff Bezos: founder public-equity operating platform plus capital allocation/private optionality
- Bernard Arnault & family: luxury brand portfolio plus family-control wealth engine
- Larry Ellison: enterprise software/database lock-in plus Oracle cloud infrastructure

If a 2026 Top 100 person does not fit an existing archetype, stop and create a new archetype template before generating more reports.

## Dashboard

The existing Streamlit dashboard remains oriented around generated processed outputs:

```powershell
streamlit run app.py
```

## Important Limitations

Forbes net worth values are annual estimates, not audited personal balance sheets. Some rows are published as “& family”; this project does not split family fortunes. Growth metrics are calculated only where sufficient history exists and should not be described as true exponential growth unless data coverage and fit quality support that interpretation.
