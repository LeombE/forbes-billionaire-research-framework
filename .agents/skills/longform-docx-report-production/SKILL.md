---
name: longform-docx-report-production
description: Use when generating detailed Microsoft Word reports for individual billionaire business analysis with charts, citations, appendices, and data-quality notes.
---

# Longform DOCX Report Production Skill

Use this skill to create `/reports/people/<year>/{rank}_{slug}_business_analysis_<year>_<variant>.docx` files for target-year reports. Legacy 2025 files may remain in `/reports/people/`.

## Report requirements
- Use the person report template in `prompts/01_PERSON_REPORT_PROMPT_TEMPLATE.md`.
- Include tables for Forbes snapshot, growth metrics, wealth equation, source list, and data-quality flags.
- Include charts generated from project data where available.
- Add captions, source notes, and alt text for figures.
- Use a source appendix with citation IDs linked to `source_citations.csv`.
- Do not pad. Expand only where evidence supports depth.

## Editorial standard
The report should read like a professional business strategy memo: specific, evidence-backed, analytical, and useful to entrepreneurs/investors/operators.
