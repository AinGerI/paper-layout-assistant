# Profile sources and boundaries

The engine accepts schema version 1 YAML, with UTF-8 or UTF-8 BOM encoding.
Keep custom profiles in the user's document project, e.g. `profiles/华为杯.yaml`,
`profiles/学校毕业论文.yaml`, `profiles/某期刊.yaml`, `profiles/导师要求.yaml`.
Do not label the demo as an official format.

## Natural-language requirements

Codex translates explicit requirements into the schema illustrated by `profiles/profile-template.yaml`.
Leave unknown settings null or disable their section. Record the source in `source_requirements`.
Validate using the installed Python: `scripts/check_profile.py <profile.yaml>`.
No external LLM API or new dependency is required; this is an agent workflow, not an autonomous parser.

## Official Word template

Run `scripts/extract_profile.py <template.docx> --output <new-profile.yaml>`.
The extraction reads built-in styles and the first section's margins. It does not infer all direct
formatting, table layouts, images, section-specific rules, custom styles or page numbering.
Review `extraction_notes`, reconcile with the written specification and run check_profile.py.

## PDF / Word specification interface

Input: specification file plus an optional official Word template.
Intermediate: a requirements list with source page/section, exact value and unknown items.
Output: schema version 1 profile and a list of unsupported/ambiguous requirements.
For Word, inspect_word.py can expose paragraph text; for PDF, use available PDF reading/OCR tools.
Codex then maps explicit requirements to the schema and validates the profile.
An autonomous PDF/OCR-to-profile parser is not implemented. Never pretend the engine has read a
specification merely because a YAML file exists. Unsupported rules require explicit manual review.

## Table properties

- `table.alignment`: whole-table position.
- `table.paragraph.alignment`: horizontal text alignment inside cells.
- `table.cell.vertical_alignment`: vertical cell alignment.
- `table.paragraph`: paragraph spacing, line spacing and indentation.
- `table.row`: height rule, height, pagination and repeating header.
- `table.cell.margin_*_pt`: internal cell padding.
- `table.border`: border structure and line widths.

OMath ranges keep their native font and structure. Fonts apply only to text outside OMath ranges.
Formula paragraphs can use `equation_block.paragraph`; exact spacing should be used only when
explicitly required. Complex merged tables fail visibly when Word rejects an unsupported operation.
