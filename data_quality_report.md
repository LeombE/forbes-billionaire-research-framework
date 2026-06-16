# Data Quality Report

This file documents project-level data-quality gates for the multi-year Forbes Top 100 system. Generated year-specific quality reports may also appear under `reports/annual/<year>/`.

## Current Status

- 2025 outputs exist and are preserved as historical baseline artifacts.
- 2026 infrastructure exists, but canonical 2026 annual data has not yet been collected in this run.
- 2026 manual import is required unless official Forbes annual data can be accessed in a compliant way.
- Current Python defaults should point to 2026; 2025 flat outputs are protected as legacy artifacts.
- GitHub-safe mode excludes local Forbes raw data, populated annual CSVs, processed outputs, Excel workbooks, generated charts, and DOCX reports from publication.

## Required Annual Gates

- Top 100 file contains exactly 100 unique people.
- Duplicate annual ranks are allowed only if Forbes uses ties and ordered position remains unique.
- Every person has annual target-year net worth from the canonical annual source or manual annual source note.
- Every row has a Forbes annual source URL, profile URL, source file, or manual source note.
- No Forbes Real-Time values overwrite annual rank or annual net worth.
- Growth metrics are blank where fewer than three valid annual observations exist.
- Source citations cover non-derived fields used in reports.
- Generated report filenames include three-digit rank, lowercase slug, target year, and variant.
- Year-specific outputs must not overwrite older years.

## 2026 Manual Import Status

Created templates:

- `templates/manual_import_top100_2026.csv`
- `templates/manual_import_wealth_history_2026.csv`
- `templates/manual_import_source_citations_2026.csv`
- `templates/manual_import_person_evidence_pack_2026.csv`

These templates are intentionally blank. Filling them is the next required step if official annual data cannot be collected programmatically.

## Known Limitations

Forbes estimates are annual snapshots, not audited personal balance sheets. Private-company valuations, ownership stakes, trusts, taxes, debt, liquidity discounts, and family ownership structures require separate evidence before final individual reports.

## GitHub-Safe Data Gates

- Public manual-import templates must remain blank.
- Public sample files must use fake people, fake companies, fake URLs, and explicit `SYNTHETIC SAMPLE` notes.
- No public CSV sample may include real Forbes Top 100 names or annual net-worth values.
- Generated outputs should be validated locally but not committed unless source rights are explicitly cleared.
