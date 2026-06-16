# Wealth Engine and Business Empire Classification Rubric

Classify each billionaire by the dominant current mechanism that explains 2025 net worth.

## Primary category selection
Choose the category that explains the majority of 2025 net worth, not necessarily the origin story.

Examples of decision rules:
- Founder/operator public equity: Most wealth from founder stake in a listed company where market capitalization drives net worth.
- Founder/operator private company: Most wealth from founder stake in a private operating company or private holding company.
- Early employee/executive equity: Significant wealth from joining early or leading, but not founding, a high-value company.
- Investor/capital allocator: Wealth primarily from investing, portfolio construction, insurance float, private equity, hedge funds, or capital allocation.
- Inherited/family-controlled business: Wealth primarily inherited or controlled through family ownership structures.
- Luxury/retail brand ownership: Wealth from branded consumer/luxury/retail groups where brand equity, pricing power, and distribution dominate.
- Technology/platform monopoly/network effects: Wealth from platform economics, network effects, software scale, data, or ecosystem control.
- Real estate/land/infrastructure: Wealth from property, land banks, infrastructure concessions, toll assets, construction, or rent extraction.
- Commodities/energy/resources: Wealth from oil, gas, mining, metals, chemicals, energy infrastructure, or resource cycles.
- Diversified holding company: Wealth from multi-sector holding company where no single operating category dominates.
- Other/unclear: Evidence is insufficient or the fortune does not fit categories.

## Confidence score
- High: Strong primary-source evidence identifies the core asset and ownership mechanism.
- Medium: Reliable secondary sources identify the asset, but ownership/valuation details are incomplete.
- Low: Source of wealth is broad or conflicting; key ownership/valuation data is missing.

## Required output columns
- wealth_engine_category
- secondary_wealth_engines
- classification_confidence
- public_equity_dependency_flag
- key_asset_or_company
- evidence_summary
- source_ids
