# Subagent Orchestration Prompt

Use this from the project root after installing the prompt pack.

Spawn specialized subagents and wait for all of them before making final edits.

Please spawn these agents in parallel:

1. `source_verification_researcher`
   - Audit official Forbes annual-list accessibility.
   - Identify canonical sources and profile URL strategy.
   - Define what must go into manual import if Forbes blocks automated access.
   - Return a source policy and citation schema.

2. `forbes_data_engineer`
   - Inspect repo layout and current files.
   - Design/repair source modules, data schemas, validation, Excel workbook, Streamlit dashboard, and tests.
   - Return implementation plan with exact files to change.

3. `wealth_history_modeler`
   - Implement and test CAGR, log-linear growth slope, exponential-fit R², doubling time, peak net worth, max gain/loss, annual volatility, largest drawdown, completeness score.
   - Return edge cases and test cases.

4. `financial_statement_analyst`
   - Define data extraction patterns for public companies, private companies, family-controlled businesses, investment portfolios, commodities, real estate, and holding companies.
   - Return a valuation bridge template.

5. `business_strategy_analyst`
   - Build the first-principles framework for each person.
   - Create the wealth engine classification rubric and cross-industry insight taxonomy.

6. `docx_report_editor`
   - Design the Word report template, section hierarchy, citation appendix format, chart insertion plan, and per-person output filename convention.

7. `qa_auditor`
   - Define hard QA gates for hallucination, missing citations, rank uniqueness, 100-person count, invalid metrics, and manual-import transparency.

After all subagents return, synthesize their outputs, implement the project files, run tests, run the pipeline, generate reports, and explain remaining manual tasks. Do not fabricate data.
