# Format modes and acceptance checks

Use only the sections relevant to the user's deliverable. None defines a mandatory font, page size, indentation, or word count.

## Template and style inspection

Read relevant headings/body/captions/references and their style chains, document defaults, numbering definitions, and section properties. Sample non-TOC body content. Reusing a paragraph's direct properties alone does not reproduce its inherited style.

Inspect `w:ind` character and distance attributes, `w:spacing`, `w:snapToGrid`, section `w:docGrid`, and keep rules when indentation or vertical spacing looks wrong. Clear/override conflicting attributes together for the exact paragraphs being changed.

- Ordinary prose can inherit the agreed first-line indent.
- Numbered paragraphs may be flush left or hanging, depending on the request. Do not assume a bibliography setting belongs on numbered prose.
- Bibliography entries may legitimately use hanging indents even when ordinary numbered paragraphs do not.
- Heading auto-numbering can restart or duplicate typed numbers in a standalone extract. Preserve a functioning sequence or disable the unwanted auto-numbering locally and retain the requested visible number.
- A standalone document should not inherit an unwanted cover, blank page, page-number restart at zero, or unused instructional text.

## Three-line tables and references

Three-line tables have a top border, a separator below the header, and a bottom border. Suppress other table/cell/style borders; a table-level setting alone may not override cell or style borders.

Use content-appropriate widths. Keep symbols, definitions, and units in their corresponding columns. Match any requested equation representation inside cells. Preserve table notes and distinguish a formula-containing table from an equation's layout frame.

Do not split a row unless justified. Repeat headers for long tables. For a short table, keeping its caption and rows together may prevent broken layout, but do not force a table taller than the page to remain unbroken. Check every rendered row: the XML can contain all values while an export clips or omits a continuation.

For bibliography-only requests, reuse the applicable reference formatting and preserve entry data/order unless a standardization or fact-check task is also authorized. Remove unintended leading whitespace without destroying a deliberate hanging indent. Do not automatically abbreviate authors, add identifiers, or change publication metadata.

## Figures and assets

Resolve the exact referenced figure, not merely a similarly named file. Extract Word images through the relationship mapping. Preserve captions and check Word crop/transform settings if they affect the intended appearance.

Use the available resolution and preserve aspect ratio. For TIFF/transparent image conversion, retain alpha where supported or composite onto the intended page background. Inspect the result for black backgrounds and clipped labels. Retain originals when reproducibility or recovery matters; do not duplicate all assets by default for a trivial task.

Missing artwork remains a disclosed placeholder. A new scientific figure or redesigned diagram requires authorization beyond faithful conversion.

## Length comparisons

Compare under the same format basis. Preserve version-specific bold lead-ins or paragraph structures only when these are part of the comparison, and state their contribution to length.

Report counts with a defined convention (for example, non-whitespace characters including punctuation, excluding headings). Separately report rendered pages and approximate last-page occupancy. Two versions can each occupy two pages but use very different amounts of space.

Do not estimate final pagination from text counts alone. A page-limit request does not silently authorize smaller fonts or removal of scientific content. If exact text and fixed formatting make the limit infeasible, deliver the faithful result, identify the conflict, and obtain the specific tradeoff needed for further edits.

For an appended comparison, preserve the prior document's contents and add clearly labeled versions. For a standalone request, include only the requested section and necessary heading/notes.

## Replacement-fragment documents

Use this pattern without reconstructing the whole paper:

1. Section and precise replacement boundary.
2. A short, verified nearby source phrase for location, styled distinctly.
3. Final replacement text/code with no editorial comments mixed into the copyable content.
4. A short preservation note only where ambiguity could cause an unintended edit.

For a one-line equation change, identify which existing line is replaced and whether outer delimiters/environments belong to the copy. For a paragraph replacement, give the full replacement paragraph. Keep the locator and instruction with enough replacement content to prevent an orphaned heading.

For coupled technical revisions, update the introduction, formula, definition, and relevant constraint prose together. Do not globally replace a result metric merely because a similarly named constraint changed. Keep pending recomputation or data synchronization in a change record, not as an invented validation result.

## Verification frontier

For faithful conversion, compare ordered content rather than only searching for a few expected phrases. Include table cells, captions, inline math, and notes. For edited prose, compare against the accepted rewrite and assert any retained numerical claims.

For scoped existing-document changes, validate untouched content and relevant package parts directly. XML-local edits can preserve unrelated parts byte-for-byte; a library save can reserialize them, so choose a proportionate content/style check instead of claiming byte identity without evidence.

Render the saved final artifact and check:
- Correct heading hierarchy and no duplicate numbering.
- Readable body spacing, no stretched short lines, and no stacked indentation.
- Complete tables, attached captions, and no missing equation rows.
- Readable glyphs and units; distinguish a renderer's missing fonts from missing source text.
- Correct page starts, page numbers, and no unexpected blank pages.

If a render fails or produces missing text, resolve the environment mismatch before changing the document's specified fonts or claiming completion. Keep temporary renderer font configuration task-local where possible. After a layout-affecting change, refresh the preview; identify the rendering engine's boundary if it differs from the user's final application.
