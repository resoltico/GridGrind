---
afad: "4.0"
version: "0.63.0"
domain: STRUCTURED_FEATURE_MUTATIONS
updated: "2026-05-01"
route:
  keywords: [gridgrind, structured feature mutations, style, validation, table, pivot, named-range, append-row]
  questions: ["where is the structured feature mutation reference", "how are structured feature mutations split in gridgrind", "where do i find table or validation mutations in gridgrind"]
---

# Structured Feature Mutation Reference

**Purpose**: Landing page for style, validation, table, pivot, named-range, and execution-policy
mutations.
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md),
[CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

The long-form structured-feature reference is split so style/validation rules and structured-data
flows stay easier to audit independently.

| Detailed reference | Owns |
|:-------------------|:-----|
| [STYLE_AND_VALIDATION_MUTATIONS.md](./STYLE_AND_VALIDATION_MUTATIONS.md) | `APPLY_STYLE`, data validations, and conditional formatting |
| [STRUCTURED_DATA_MUTATIONS.md](./STRUCTURED_DATA_MUTATIONS.md) | autofilters, tables, pivot tables, `APPEND_ROW`, `AUTO_SIZE_COLUMNS`, `execution.calculation`, and named ranges |
