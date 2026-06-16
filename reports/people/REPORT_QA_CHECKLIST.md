# Report QA Checklist

## Data Identity

- Report year matches the canonical dataset year.
- Rank and name match `top100_<year>.csv`.
- Filename uses three-digit rank, lowercase slug, year, and variant.
- Forbes Real-Time data is absent from annual canonical fields.

## Evidence

- Every body citation key appears in the evidence appendix.
- No raw URLs appear in body text before the appendix.
- Annual Forbes fields cite annual sources or manual annual source notes.
- Business, financial, ownership, valuation, and risk claims cite primary or high-quality sources.
- Evidence gaps remain visible where facts are unavailable.

## Formatting And Files

- DOCX is a valid ZIP and opens with `python-docx`.
- Referenced charts exist under `reports/charts/<year>/`.
- Embedded images are present when charts are referenced.
- Appendix tables are readable.
- Data limitations and confidence notes remain in the document.
- Generated DOCX reports are local research outputs and should not be committed to a GitHub-safe public release.

## Leakage Prevention

Scan for unintended prior-template terms from 2025 references:

- Elon / Musk / Tesla / SpaceX
- Mark / Zuckerberg / Meta
- Jeff / Bezos / Amazon / AWS
- Bernard / Arnault / LVMH / Dior
- Larry / Ellison / Oracle

Intentional peer-comparison mentions are allowed only in the comparable patterns section.
