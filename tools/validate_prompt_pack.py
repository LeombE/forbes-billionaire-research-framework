from pathlib import Path

required = [
    'AGENTS.md',
    '.codex/config.toml',
    '.codex/agents/source_verification_researcher.toml',
    '.codex/agents/forbes_data_engineer.toml',
    '.codex/agents/wealth_history_modeler.toml',
    '.codex/agents/financial_statement_analyst.toml',
    '.codex/agents/business_strategy_analyst.toml',
    '.codex/agents/docx_report_editor.toml',
    '.codex/agents/qa_auditor.toml',
    '.agents/skills/forbes-billionaire-data-pipeline/SKILL.md',
    '.agents/skills/billionaire-business-empire-analysis/SKILL.md',
    '.agents/skills/wealth-growth-modeling/SKILL.md',
    '.agents/skills/longform-docx-report-production/SKILL.md',
    '.agents/skills/source-quality-and-legal-compliance/SKILL.md',
    'prompts/00_MASTER_CODEX_PROMPT.md',
    'prompts/01_PERSON_REPORT_PROMPT_TEMPLATE.md',
    'prompts/02_SUBAGENT_ORCHESTRATION_PROMPT.md',
    'prompts/03_IMAGE_AND_VISUAL_POLICY.md',
    'templates/manual_import_top100_2025_template.csv',
    'templates/wealth_engine_rubric.md',
]

missing = [p for p in required if not Path(p).exists()]
if missing:
    print('Missing files:')
    for p in missing:
        print(f' - {p}')
    raise SystemExit(1)
print('Prompt pack validation passed.')
