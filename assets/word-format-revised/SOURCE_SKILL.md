---
name: word-format
description: Analyze, reformat, and validate Microsoft Word .docx academic documents on Windows using the installed desktop Microsoft Word. Use when the user wants Word formatting standardized or converted to a competition, journal, thesis, school, or custom format without switching to LaTeX.
---

Use this skill for Microsoft Word `.docx` formatting on Windows. The engine controls the user's installed Microsoft Word through `pywin32` COM automation. Never overwrite the source document.

## Core rule

Separate **format requirements** from the **formatting engine**. The engine stays the same. Each competition, journal, thesis guide, or custom request becomes a YAML profile under `profiles/`.

## Input priority

When building a profile, prefer these sources in order:

1. Explicit formatting requirements supplied by the user.
2. An official Word template. Run `scripts/extract_profile.py` to create a starter profile, then reconcile it against written requirements.
3. A PDF/Word formatting guide. Read the guide and translate explicit requirements into a profile.
4. Natural-language requirements from the user.

Do not invent missing format requirements. If a requirement is unknown, keep the corresponding setting disabled or null so the existing document formatting is retained.

## Workflow

1. Run `scripts/inspect_word.py` on the source `.docx` before changing it.
2. Review the generated structure report. Resolve obvious semantic misclassification with `detection.paragraph_overrides` or `detection.regex_overrides` in the profile.
3. Create or update the requested YAML profile.
4. Run `scripts/format_word.py` with an explicit output path. This script copies the source first and formats only the copy.
5. Read the generated formatting report.
6. Always run `scripts/validate_word.py` to compare the saved source and formatted files.
7. Tell the user which items still need manual Word review, especially floating shapes, merged/irregular tables, unusual equation numbering, or layout-sensitive pagination.

## Safety requirements

- Never overwrite the original `.docx`.
- Never rewrite body text, numbers, references, equations, captions, or table contents merely to change formatting.
- Do not convert Word equations to images.
- Do not automatically update fields, cross-references, bibliography fields, or TOC unless the user explicitly requests it.
- Formula-heavy paragraphs should not receive fixed line spacing by default. Prefer `at_least` spacing unless the official format explicitly requires exact spacing.
- Table row height should default to `auto`; exact height is allowed only when explicitly requested.
- Floating Word shapes are layout-sensitive. Resize/reposition them only when explicitly requested.
- If exact body text or object counts change after formatting, surface the warning prominently and do not claim the file is ready to submit. Integrity checks cover the main document story, not all header/footer/textbox stories.

## Commands

Resolve this SKILL.md's directory first and use its `.venv/Scripts/python.exe` and scripts by absolute path. Do not depend on the user's current directory. Keep generated profiles and reports beside the user's task files; bundled profiles are examples. Word COM requires access to the interactive Windows desktop; if sandbox activation fails, request the narrow desktop execution permission and rerun the same command.

Unknown/null font and paragraph settings retain current values. Invalid profile keys are rejected before Word starts. Existing outputs and colliding report/input paths are refused; use another output name. Commands do not overwrite previous results. A failed run may leave an incomplete output copy and failure reports; these are not deliverables.

For profile creation from natural language, templates or PDF/Word specifications, read `references/profile-sources.md`.

Analyze:

```powershell
.\.venv\Scripts\python.exe .\scripts\inspect_word.py "paper.docx" --profile ".\profiles\target.yaml"
```

Extract a starter profile from a Word template:

```powershell
.\.venv\Scripts\python.exe .\scripts\extract_profile.py "official-template.docx" --output ".\profiles\target.yaml"
```

Format:

```powershell
.\.venv\Scripts\python.exe .\scripts\format_word.py "paper.docx" --profile ".\profiles\target.yaml" --output "paper_formatted.docx"
```

Validate integrity:

```powershell
.\.venv\Scripts\python.exe .\scripts\validate_word.py "paper.docx" "paper_formatted.docx"
```

## Profile semantics

Main paragraph roles are `title`, `heading_1`, `heading_2`, `heading_3`, `body`, `equation_block`, `figure_paragraph`, `figure_caption`, `table_caption`, `abstract_heading`, `keywords`, and `references_heading`.

Use points for font size/paragraph spacing and centimeters for page/image dimensions. Chinese font size conversions and profile guidance are in `references/profile-guide.md`.

## Completion standard

A successful run must produce:

1. A new formatted `.docx`.
2. A Markdown formatting report.
3. A JSON machine-readable report.
4. A concise list of items that still need manual inspection.

A file is not considered fully verified merely because Word saved it successfully.

## 本机模板库与中文文件管理

用户使用本机模板库或按比赛类别排版时，先读取 [本机中文工作区](references/local-workspace.md)。优先使用中文目录和中文交付说明；程序内部文件名保持稳定。

