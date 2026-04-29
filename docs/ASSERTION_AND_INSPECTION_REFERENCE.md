---
afad: "3.5"
version: "0.61.0"
domain: ASSERTION_INSPECTION_REFERENCE
updated: "2026-04-25"
route:
  keywords: [gridgrind, assertions, inspections, analysis, get-cells, charts, workbook-health]
  questions: ["where is the assertion and inspection reference", "how are assertions and inspections split in gridgrind", "where do i find workbook health queries in gridgrind"]
---

# Assertion And Inspection Reference

**Purpose**: Landing page for ordered assertions, factual inspection queries, and finding-bearing
analysis queries.
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md),
[CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md), and
[STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md)

Assertions and inspections are still ordered `steps[]` entries. The detail is now split so
assertion families, factual read surfaces, and analysis/query families can be reviewed separately.

| Detailed reference | Owns |
|:-------------------|:-----|
| [ASSERTIONS.md](./ASSERTIONS.md) | ordered assertion families, shapes, and failure semantics |
| [WORKBOOK_AND_CELL_INSPECTIONS.md](./WORKBOOK_AND_CELL_INSPECTIONS.md) | workbook-core, sheet-core, cell, window, hyperlink, and comment inspections |
| [DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md) | drawing, layout, validation, table, pivot, schema, and structural inspections |
| [ANALYSIS_QUERIES.md](./ANALYSIS_QUERIES.md) | finding-bearing analysis queries, formula health, workbook health, and aggregate findings |
