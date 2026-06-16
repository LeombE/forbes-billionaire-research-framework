# Archetype Routing Guide

This guide routes annual Forbes Top 100 enriched reports to the closest analytical template. It is multi-year infrastructure; the existing 2025 rank 1-5 reports are references, not 2026 evidence.

## Current Reference Archetypes

| Archetype | Use when | 2025 reference |
|---|---|---|
| Founder public equity plus private frontier optionality | Public founder equity dominates, with separate private optionality. | Elon Musk |
| Platform/network-effects advertising engine | User network, data/recommendation loop, ads or platform monetization drives value. | Mark Zuckerberg |
| Founder public-equity operating platform plus capital allocation/private optionality | Listed operating platform dominates, with private assets treated conservatively. | Jeff Bezos |
| Luxury brand portfolio plus family-control wealth engine | Family-controlled brand portfolio, pricing power, selective distribution, and voting control drive value. | Bernard Arnault & family |
| Enterprise software/database lock-in plus cloud infrastructure | Installed-base software economics, cloud migration, support renewals, and founder equity drive value. | Larry Ellison |

## Additional Archetypes To Add When Needed

- Investor/capital allocator
- Inherited/family-controlled retail
- Commodities/energy/resources
- Real estate/land/infrastructure
- Diversified holding company
- Early employee/executive equity
- Other/unclear

## 2026 Routing Rules

- Use `data/interim/2026/archetype_routing_table_2026.csv` for route metadata.
- Treat 2025 DOCX reports as `template_reference_only=True`.
- Do not copy person-specific 2025 prose into 2026 reports.
- If a 2026 person does not fit a route, stop and create a new archetype template before continuing the batch.
- Person-specific company, ownership, valuation, and risk claims must come from 2026 evidence packs or current primary sources.

## Evidence Tables

- `data/processed/<year>/source_citations_<year>.csv` supports annual Forbes/project fields and derived metrics.
- `data/interim/<year>/enriched_evidence_registry_<year>.csv` supports report-level business evidence.
- A 2025 registry row cannot satisfy a 2026 report validation gate.

## GitHub-Safe Publication

The routing guide can be public, but enriched evidence registries and generated report outputs should remain local unless source redistribution rights are confirmed. Public examples should use synthetic sample evidence only.
