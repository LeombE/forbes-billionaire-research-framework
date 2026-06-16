# Citation Policy

Every material factual claim in a processed table or individual report must trace to a source row, source file, or explicit source note.

## Required Citation Fields

- `source_id`
- `person_name`
- `field_name`
- `claim_supported`
- `source_title`
- `source_url` or `source_file`
- `publisher`
- `author` when available
- `publication_date`
- `accessed_at`
- `evidence_note`
- `reliability_tier`
- `limitations`

## Citation Rules

- Forbes annual list rows support annual ranking fields only.
- Company filings, annual reports, proxy statements, and investor-relations documents support financial-statement and ownership claims.
- Secondary sources can support background context, not unsupported precise ownership or valuation bridges.
- Body text in DOCX reports should use short citation keys; full URLs belong in the evidence appendix.
- Any body citation key must appear in the appendix.
- Public sample citations must use synthetic locators such as `https://example.com/...` and must be marked as fake demonstration data.

## Unsupported Claims

Unsupported ownership stakes, private valuations, taxes, debt, trusts, market values, or personal liquidity claims must be removed or labelled as unknown.

## GitHub-Safe Citation Boundary

`source_citations.csv` and year-specific source citation outputs are local generated research artifacts. Public repositories should include citation schema templates and synthetic examples only, not populated Forbes-derived citation tables.
