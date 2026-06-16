# Install Guide — Forbes Billionaire Codex Prompt Pack

This pack is designed for your existing `fobes top 100` / Forbes Top 100 project folder.

## What to copy
Copy everything inside `copy_into_repo/` into the root of your project repository, next to files like `app.py`, `README.md`, `requirements.txt`, `/src`, `/data`, and `/reports`.

After copying, your repo should contain:

```text
AGENTS.md
.codex/config.toml
.codex/agents/*.toml
.agents/skills/*/SKILL.md
prompts/*.md
templates/*.md
templates/manual_import_top100_2025_template.csv
tools/validate_prompt_pack.py
```

## Windows PowerShell copy example
From the folder where you extracted this pack:

```powershell
$Repo = "C:\Users\<YOU>\OneDrive\Leo - Personal\Documents\fobes top 100"
Copy-Item -Recurse -Force .\copy_into_repo\* $Repo
```

Then open a terminal in the project root and run:

```powershell
codex --ask-for-approval never "Summarize the current instructions and list available skills/subagents for this Forbes project."
```

## Verify the pack
Inside the project root:

```powershell
python .\tools\validate_prompt_pack.py
```

Expected result: it prints `Prompt pack validation passed.`

## Suggested execution sequence
1. Paste `prompts/00_MASTER_CODEX_PROMPT.md` into Codex from the project root.
2. Ask Codex to spawn the subagents listed in `.codex/agents/`.
3. Run the data pipeline first. Do not start 100 Word reports until `top100_2025.csv`, `source_citations.csv`, and growth metrics pass validation.
4. Generate reports in batches of 5–10 people to avoid context blow-up and to improve citation QA.
5. Run QA after every batch.

## Important limitation
A prompt pack cannot bypass Forbes access controls. If Forbes blocks automated retrieval or requires a subscription, use the manual-import template and document the source fields supplied by the human researcher.
