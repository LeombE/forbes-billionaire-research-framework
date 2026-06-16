# Batch Generation Protocol

Do not generate all 100 reports in one run. Generate one report, validate it, then move to a small batch only after source coverage is adequate.

## Generate One 2026 Report

```powershell
.\.venv\Scripts\python.exe -m src.enriched_reports --year 2026 --rank 1 --variant enriched_draft
```

Expected output:

```text
reports/people/2026/001_<name_slug>_business_analysis_2026_enriched_draft.docx
```

## Small Batch Rule

After rank 1 passes validation, generate at most ranks 2-5 or another small explicitly requested range. Stop on the first validation failure.

## Stop Conditions

- No canonical annual 2026 row exists.
- Forbes Real-Time data is being used as annual rank or annual net worth.
- No matching archetype exists.
- Evidence pack lacks primary or official support for ownership, financial, valuation, or risk claims.
- Body citation keys do not appear in the appendix.
- Previous-person template text leaks into the report.
- Referenced chart files do not exist.
- The target filename already exists and overwrite was not explicitly approved.

## Versioning

Use new filenames for upgraded reports:

- `_enriched_draft.docx`
- `_enriched_v2.docx`
- `_final_review_ready.docx`

Do not overwrite previous versions unless explicitly approved.

## GitHub-Safe Release Boundary

Generated report batches, chart PNGs, and enriched evidence registries stay local. Public releases should include report-generation code, QA checklists, templates, and synthetic samples, not generated billionaire reports.
