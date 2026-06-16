# Contributing

## Local Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pytest
```

## Data Contributions

- Use official annual Forbes list data for canonical annual fields.
- Do not submit copied raw Forbes pages or paywalled content.
- Use manual-import templates when automated access is blocked or unclear.
- Include citations for every manually entered non-derived field.
- Do not submit populated Forbes annual CSVs, raw Forbes JSON, processed outputs, Excel workbooks, charts, or DOCX reports to the public repository.
- Use `samples/` only for fake demonstration rows; every sample row must be clearly marked synthetic.

## Report Contributions

- Generate one person report first; do not batch-generate all 100.
- Keep evidence gaps visible.
- Do not overwrite previous DOCX variants without explicit approval.
- Validate citation keys, evidence appendix, chart files, filename convention, and previous-person leakage.

## Pull Request Checklist

- Tests pass.
- New year outputs do not overwrite older years.
- No secrets, cache files, `.venv`, `__pycache__`, logs, or proprietary raw files are included.
- No Forbes-derived local data outputs or generated billionaire reports are included.
- Blank templates remain blank; populated manual-import files stay private.
- Documentation explains limitations and source status.
