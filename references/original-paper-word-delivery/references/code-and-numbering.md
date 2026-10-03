# Code appendices, captions, and numbering

Read only for code-containing papers or numbering/layout issues. These notes do not define a universal font, highlighting theme, or page format.

## Code appendices

Separate an appendix excerpt from a runnable program package. An excerpt can contain core implementation only; a runnable package must include its actual modules and declared inputs. Do not claim a code-only archive is self-contained when required data are absent.

Preserve source code semantics, indentation, and meaningful comments. For a formatting-only request, do not refactor algorithms or silently remove comments. When comment reduction is authorized, keep explanations of non-obvious formulas, units, boundaries, and decisions; omit line-by-line narration.

Select typeface, size, spacing, borders, and highlighting from the requested template and reading context. A monospaced face often helps code alignment, but a user's chosen Times New Roman setting is valid; inspect the resulting alignment instead of forcing a font preference. Compact spacing must not clip glyphs or merge lines visually.

If highlighting is requested, use an available highlighter or converter that retains editable text when the target supports it. A colored screenshot is not editable code. Check long lines, tabs/spaces, page breaks, and contrast in the actual document. Do not require Pandoc, a named theme, or a particular template repository when another current tool is adequate.

If a new snippet has to remain executable, keep syntactic validity and its required context. Do not insert typographic quotes, soft hyphens, or mathematical glyph substitutions into program text.

## Captions and cross-references

Reuse existing numbering and caption conventions. For newly structured documents, native fields and heading-linked numbering can improve maintainability when supported by the chosen workflow; do not convert all manual numbering during an unrelated narrow edit.

Check that a numbered heading is a real paragraph rather than a soft line break inside the preceding heading. Bind chapter-based numbering to the intended heading level, not to a hardcoded three-level hierarchy.

Caption text, source order, displayed numbers, and cross-references should agree after insertion, deletion, or movement. Updating fields is application-specific; use the actual application/version workflow rather than assuming a universal shortcut.

MathType numbering and visible TeX delimiters are separate mechanisms. Toggle TeX alone should not be represented as creating an automatic equation-numbering system. Preserve or remove numbers according to the task and check affected references.

## Compact layout

Prefer styles, paragraph spacing, row settings, and appropriate keep rules over repeated blank paragraphs. Resolve character-based and point-based indentation together when they conflict.

For multiline formulas and tall glyphs, a very small fixed line height can clip content. Keep the requested visual compactness while checking the actual rendering. The historical use of 6-point spacing is an example, not a required setting.

Long tables may need repeated header rows. Set and inspect that behavior on the affected tables; do not assume a visual three-line table style alone supplies repeating headers.

Changes to code, captions, equations, and bibliography have different scopes. A local compact-code request should not reset the entire manuscript's fonts and spacing.

