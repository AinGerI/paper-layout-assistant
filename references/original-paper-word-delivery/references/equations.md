# Equations: representation, compatibility, and recovery

Read for equation-bearing Word deliverables or Word/TeX conversion. Preserve the requested mathematical content and representation; equation appearance and editability are separate acceptance criteria.

## Choose the required state

| State | What the user receives | Verification boundary |
|---|---|---|
| Visible TeX | Editable text such as `$...$` or `\[...\]` | Code/text checks and any performed syntax checks; not an embedded formula. |
| Word-native equation | OMML equation object | Inspect equation structure and rendering; not automatically a MathType OLE object. |
| MathType equation | Embedded object recognized by MathType | Inspect object identity and preview; actual editing/import requires the target application's test. |
| Image | Visual representation | Appearance only; no claim of formula editability. |

Use conversation context to interpret “MathType 可编辑.” If visible TeX was already chosen, continue that contract and name it accurately. If the user explicitly requests native MathType objects, do not silently substitute TeX or OMML; explain a genuine capability gap and seek the needed choice.

## Visible TeX delivery

- Preserve the agreed delimiters; commonly `$...$` for inline and `\[...\]` for display math. Delimiters alone do not establish support for every enclosed command.
- Remove accidental Markdown escaping, not semantic TeX backslashes. Expand custom macros only from known definitions and check equivalence.
- Store display formulas in ordinary paragraphs when no frame is requested. Unwrap only the identified equation tables/frames; never remove real data tables.
- Preserve original multiline breaks unless the user authorizes a different layout. A compatibility adjustment must not flatten a multiline formula or drop cases, bounds, subscripts, or operators.
- Remove displayed equation numbers only when requested. Preserve prose references unless their revision is authorized; flag references left without visible targets when materially confusing.
- Prevent clipping and avoid page breaks inside short display formula blocks. Let long raw code wrap visually without changing its mathematical tokens.

## Alignment: do not guess compatibility

The intended result may be “relation signs aligned, ordinary space around each sign, objective row unaffected.” Verify all of those requirements, not just that rows align.

- Standard LaTeX/MathJax rendering does not certify MathType input compatibility.
- `aligned`, optional row spacing such as `\\[6pt]`, array column specifications, and package macros can be version-dependent at import.
- Splitting a relation into separate matrix columns can create excessive gaps. Changing three columns to two does not prove that the import problem is fixed.
- MathType's own relation-alignment feature and its TeX importer are different mechanisms. Current documentation must be scoped to the actual product (desktop vs web) and version.
- Look up the intended version’s native relation-alignment and matrix-alignment controls when they are relevant. Menu names and shortcuts are product- and version-specific; do not make one desktop workflow mandatory for all users.
- Single-column structures keep each expression intact, but do not automatically guarantee relation alignment. Avoid treating them as a universally verified workaround.

After a user reports failure, record that route as failed in the current task and identify whether the problem is parsing, structure, or spacing. Do not cycle through undocumented syntax with claims of certainty. If a screenshot cannot be supplied, continue source/documentation inspection and state only the genuinely unresolved application-specific boundary.

Primary documentation pointers (recheck when making version-specific claims):
- [MathType 7 menus](https://docs.wiris.com/en_US/mathtype-7-user-interface/menus)
- [MathType 7 dialogs](https://docs.wiris.com/en_US/mathtype-7-user-interface/dialogs)
- [Translation limitations](https://docs.wiris.com/en_US/conversion-and-compatibility/some-latex-and-mathml-fails-to-translate)

## Recovering equations from Word

Text extraction alone can omit embedded MathType expressions, including small inline symbols and numeric values. Inventory `w:object` / OLE objects, `m:oMath` / OMML, and equation preview images, retaining their source order and association.

For OLE MathType objects, original equation data (often MTEF in an Equation Native stream) may allow structured recovery. A generic Office conversion can preserve only OLE wrappers or produce empty formulas; count and inspect outputs rather than trusting a successful process exit.

Use a suitable reader when available. Inspect unfamiliar tool code and outputs, keep working material task-local, and avoid installing permanent global dependencies merely for one conversion. Do not make an unverified reader a required dependency of this skill.

Verify recovered code against source previews:
- Fractions, signs, integral/sum limits, case conditions, subscripts/superscripts, accents, and units.
- Private-use glyphs, function names, special spaces, and multiline matrices/piles.
- All high-risk or repaired expressions and a representative inline sample; report the actual coverage.

A nonempty expression or successful TeX compilation is not proof of faithful recognition. Keep source equations/previews or a traceable mapping when recovery risk warrants it. Mark unresolved items rather than silently guessing, deleting them, or substituting equations from an older draft.


## Optional simple import structure

If the target importer struggles with a multiline environment, a simple single-column array is one candidate to test. For example:

~~~latex
\[
\begin{array}{l}
\min_{\mathbf{u}}\quad f(\mathbf{u})\\
\mathrm{s.t.}\quad g_r(\mathbf{u})\leq 0,\quad r=1,\ldots,m
\end{array}
\]
~~~

This is an illustrative input, not a verified conversion for every MathType version and not a requirement to rewrite all equations. It preserves rows but does not promise aligned relation signs.

A previous local workflow encountered issues with aligned environments, optional row spacing, and Chinese text in TeX input. Treat these as diagnostic leads when the same symptom appears; do not ban valid LaTeX globally. Test a representative expression in the actual target application before claiming that a compatibility problem is solved.
