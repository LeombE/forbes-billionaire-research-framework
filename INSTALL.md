# Install Guide

This project is a Python research framework for annual Forbes World's Billionaires Top 100 analysis. Public releases include code, documentation, blank manual-import templates, and synthetic sample data only.

## Requirements

- Python 3.11 or newer
- Windows PowerShell, macOS Terminal, or a compatible shell
- Source data that you are legally allowed to use

## Create a Virtual Environment

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Verify the Installation

Run the test suite:

```powershell
.\.venv\Scripts\python.exe -m pytest
```

On macOS or Linux:

```bash
python -m pytest
```

## Prepare a Target Year

Create year-specific folders and blank manual-import templates:

```powershell
.\.venv\Scripts\python.exe -m src.pipeline --year 2026 --init-year
```

Do not run collection or analysis commands until your source data is ready and your source permissions are clear.

## Manual-Import Workflow

If official annual data cannot be collected programmatically in a compliant way, populate the manual-import templates in a private local workspace:

- `templates/manual_import_top100_2026.csv`
- `templates/manual_import_wealth_history_2026.csv`
- `templates/manual_import_source_citations_2026.csv`
- `templates/manual_import_person_evidence_pack_2026.csv`

Then run:

```powershell
.\.venv\Scripts\python.exe -m src.pipeline --year 2026 --manual-import
```

## Source and License Boundary

The MIT License applies to this framework's code, documentation, blank templates, and synthetic samples. It does not include Forbes data, third-party source data, generated datasets, generated workbooks, generated charts, generated reports, or user-provided annual datasets.
