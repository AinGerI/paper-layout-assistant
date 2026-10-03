# Paper Layout Assistant · 论文排版助手

Turn manuscript Markdown into an editable Word document, then use a reusable macro in Windows Word with MathType 7 to compile equations and maintain numbering.

- Configurable typography and layout: current user requirements and supplied templates override the defaults.
- Markdown-first workflow with embedded figures, tables, equation bookmarks and paragraph checks.
- One reusable macro, including incremental conversion, chapter numbering, optional backup and in-place saving.
- Routine delivery: a Word file; include the macro only for first-time setup.

## Use

Read [SKILL.md](SKILL.md) in your agent environment. Python 3.10+ and the dependencies in `pyproject.toml` are required for conversion. Use `uv sync --locked` to prepare an environment.

```sh
uv run python scripts/markdown_to_word.py manuscript.md manuscript.docx --work-dir work/conversion --profile profiles/chinese-paper.json
uv run python scripts/audit_layout.py manuscript.docx --profile profiles/chinese-paper.json
```

The generated document contains visible TeX, not already compiled MathType objects. In Windows Word with MathType 7, import `assets/PWDMathType133Check.bas` once and run `PWD_Run133_Check`. See [macro instructions](references/macro-use.md).

The Markdown parser supports the subset described in [the input contract](references/markdown-contract.md). Existing Word files and advanced template requirements use the documented separate workflow. Paragraph checks do not emulate Word's rendering; inspect the final layout in your target environment.

## 中文

先整理 Markdown，再按本次格式要求生成 Word，最后在 Windows Word + MathType 7 中运行通用宏完成公式转换和编号。

字体、字号、对齐、缩进、间距和页边距可以配置；用户提供的新版模板与明确要求优先。日常只交付 Word，首次使用时附上宏，图片内嵌文档。

从 [SKILL.md](SKILL.md) 开始阅读。生成命令见上方；首次导入 `assets/PWDMathType133Check.bas`，以后运行 `PWD_Run133_Check` 即可。输入格式与已有 Word 的处理路径见对应说明，最终版式在目标 Word 环境检查。

## Validation / 验证

Automated converter and audit tests are in `tests/`. The macro requires Windows Word and MathType; Python tests do not execute it. No MathType software or Microsoft Word binaries are distributed here.

转换与检查工具的自动测试位于 `tests/`。宏需要 Windows Word 和 MathType，Python 测试不会运行宏。本项目不分发 MathType 软件或 Microsoft Word 程序。
