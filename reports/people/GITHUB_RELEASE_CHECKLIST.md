# GitHub Release Checklist

Use this checklist before publishing the repository publicly.

## Include

- Source code under `src/`
- Tests under `tests/`
- Blank manual-import templates under `templates/`
- Synthetic fake-only samples under `samples/`
- Project documentation, methodology, source policy, citation policy, and QA guides

## Exclude

- Forbes raw JSON/API caches
- Populated manual-import CSVs
- Processed Forbes-derived CSVs
- Excel workbooks
- Enriched evidence registries and local evidence packs
- Generated chart PNGs
- Generated billionaire DOCX reports
- Streamlit logs, caches, virtual environments, and temporary files

## Required Checks

```powershell
.\.venv\Scripts\python.exe -m py_compile src\config.py src\pipeline.py src\fetch_forbes.py src\clean.py src\metrics.py src\citations.py src\quality.py src\reports.py src\excel.py src\enriched_reports.py
.\.venv\Scripts\python.exe -m pytest
```

## Manual Review

- Confirm `.gitignore` excludes local Forbes-derived outputs.
- Confirm manual-import templates are blank.
- Confirm `samples/` rows are synthetic and visibly marked.
- Confirm no generated DOCX, PNG, XLSX, or populated Forbes CSV is staged.
- Confirm README explains source limits and manual import workflow.
