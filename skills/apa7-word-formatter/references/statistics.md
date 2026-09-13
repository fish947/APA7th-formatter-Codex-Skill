# Statistics and equations

Read this reference whenever `_review.statistics_formula_check` detects statistical expressions, native Word math, or plain-text formula candidates.

## Official basis

- [APA Numbers and Statistics Guide](https://apastyle.apa.org/instructional-aids/numbers-statistics-guide.pdf): “Report exact p values to two or three decimals (e.g., p = .006, p = .03).”
- [APA Publication Manual, Presentation of Equations](https://www.apa.org/pubs/books/publication-manual-7th-edition-spiral): “Number all displayed equations consecutively, with the number in parentheses near the right margin of the page.”
- [APA Research Transparency Standards](https://www.apa.org/pubs/journals/resources/standards-disclosures): “Exact p values, effect sizes, and 95% confidence intervals, or an explanation of why this was not possible.”

The Publication Manual, Sections 6.32–6.48, contains the full numbers, statistics and equation rules. A university, instructor or journal may impose a different reporting requirement.

## What Python may change automatically

- Normalize safe presentation in recognized expressions without changing numeric value: `P=0.032` becomes italic *p* `= .032`; `t(28)=.75` becomes italic *t* `(28) = 0.75`.
- Use italic type for recognized Latin statistical symbols and upright type for recognized Greek symbols.
- Normalize spaces around recognized comparison operators and apply the APA leading-zero convention when the statistic's range is known.
- For a paragraph confidently classified as an independent native Word equation, remove paragraph indents and extra paragraph spacing and use double line spacing. Preserve the math XML.

Every safe text edit is recorded. If a symbol or expression is split across complex fields, text boxes or unsupported runs, preserve it and include it in review rather than rebuilding the paragraph.

## What AI must review

- Check that the statistical test type is interpreted correctly and that the reported statistic, degrees of freedom, exact *p* value, effect size and confidence interval are appropriate for that analysis.
- Treat `p = .000`, an out-of-range *p* value, unusual comparison operators and questionable decimal precision as author-review items.
- Compare values with tables and the author's analysis output when available. A pattern match does not establish statistical correctness.
- For native or plain-text equations, check variables, Greek letters, functions, operators, fences, superscripts, subscripts and sentence punctuation.
- Check displayed-equation numbering, order, right-margin placement and every in-text equation reference on the rendered pages.

## Never do automatically

- Do not round or recompute a value, change `=` to `<`, change the direction of a comparison, or alter whether a result is significant.
- Do not invent a missing statistic, degrees of freedom, effect size, confidence interval or formula term.
- Do not convert plain text or an equation image into editable Word math without an explicit, separately reviewed reconstruction request.
- Do not renumber equations automatically because existing cross-references may depend on the current numbers.

The concise feedback should report how many expressions and formulas were recognized, how many safe presentation changes were made, and how many items still need author confirmation. Do not expose the full internal issue list unless the user asks for technical detail.
