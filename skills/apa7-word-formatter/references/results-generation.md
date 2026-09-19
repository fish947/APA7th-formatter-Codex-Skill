# Results to APA tables and figures

Use this mode when the author supplies structured CSV or JSON and asks for an APA-ready results document. Do not use it to infer data from a screenshot or to recompute an analysis.

## Safety boundary

- Preserve every supplied table value as written. JSON is loaded with lexical number preservation so values such as `0.10` remain `0.10`.
- Do not run a statistical test, round a result, change significance, fill a missing statistic or interpret the scientific finding.
- Flag suspicious values such as `p = .000`; do not silently change them to another claim.
- A figure is reconstructed only from supplied plotting data. The Word preview is a high-resolution PNG for compatibility; optional SVG export is genuine geometry generated from those values.
- Render and inspect every page. Check table width, repeated headers, clipped values, figure labels, legend, notes and section orientation.

## CSV

The first row supplies column names and each later row supplies one table row:

```sh
python3 scripts/apa7_results.py results.csv \
  --output results_APA7.docx \
  --table-title "Descriptive Statistics by Condition" \
  --note "M = mean; SD = standard deviation."
```

## JSON

```json
{
  "schema_version": 1,
  "document_title": "Results Tables and Figures",
  "tables": [
    {
      "number": 1,
      "title": "Descriptive Statistics by Condition",
      "columns": ["Condition", "n", "M", "SD"],
      "rows": [["Control", "30", "4.10", "0.82"], ["Treatment", "30", "4.68", "0.77"]],
      "numeric_columns": [1, 2, 3],
      "note": "M = mean; SD = standard deviation."
    }
  ],
  "figures": [
    {
      "number": 1,
      "title": "Mean Score by Session",
      "type": "line",
      "x_label": "Session",
      "y_label": "Mean score",
      "series": [
        {"name": "Control", "x": ["1", "2", "3"], "y": ["4.10", "4.20", "4.32"]},
        {"name": "Treatment", "x": ["1", "2", "3"], "y": ["4.18", "4.45", "4.68"]}
      ],
      "alt_text": "Two lines show mean scores increasing across three sessions.",
      "note": "Error bars were not supplied and were not invented."
    }
  ]
}
```

Run:

```sh
python3 scripts/apa7_results.py results.json --output results_APA7.docx
python3 scripts/apa7_results.py results.json --output results_APA7.docx --export-svg results_svg
```

Supported figures are `bar`, `line` and `scatter`. All line/bar series must share the same x labels. Use strings for reported table values when exact displayed precision matters.

## Official basis

- [APA Table Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/tables): “Give each table a brief but descriptive title, and capitalize the table title in italic title case.”
- [APA Figure Setup](https://apastyle.apa.org/style-grammar-guidelines/tables-figures/figures): “Give each figure a brief but descriptive title, and capitalize the figure title in italic title case.”
- [APA Numbers and Statistics Guide](https://apastyle.apa.org/instructional-aids/numbers-statistics-guide.pdf): “Report exact p values to two or three decimals (e.g., p = .006, p = .03).”
