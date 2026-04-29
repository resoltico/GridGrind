---
afad: "3.5"
version: "0.61.0"
domain: WORKBOOK_LAYOUT_MUTATIONS
updated: "2026-04-25"
route:
  keywords: [gridgrind, workbook mutations, sheet mutations, layout, panes, print-layout, structure]
  questions: ["where is the workbook and layout mutation reference", "how are workbook mutations split in gridgrind", "where do i find sheet and layout steps in gridgrind"]
---

# Workbook, Sheet, And Layout Mutation Reference

**Purpose**: Landing page for workbook-core, sheet, layout, and print-surface mutations.
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md),
[STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

The long-form workbook and layout reference is split so sheet/state flows and structural/layout
flows can drift less easily and stay easier to navigate.

| Detailed reference | Owns |
|:-------------------|:-----|
| [WORKBOOK_AND_SHEET_MUTATIONS.md](./WORKBOOK_AND_SHEET_MUTATIONS.md) | `ENSURE_SHEET`, sheet lifecycle, visibility, protection, workbook protection, and `IMPORT_CUSTOM_XML_MAPPING` |
| [LAYOUT_AND_STRUCTURE_MUTATIONS.md](./LAYOUT_AND_STRUCTURE_MUTATIONS.md) | merges, row and column structure, visibility, grouping, panes, zoom, presentation, and print layout |
