---
afad: "3.5"
version: "0.58.0"
domain: CELL_DRAWING_MUTATIONS
updated: "2026-04-23"
route:
  keywords: [gridgrind, cell mutations, drawing mutations, hyperlink, comment, picture, chart, signature-line]
  questions: ["where is the cell and drawing mutation reference", "how are cell and drawing mutations split in gridgrind", "where do i find chart or comment mutations in gridgrind"]
---

# Cell And Drawing Mutation Reference

**Purpose**: Landing page for cell-value, link/comment, and drawing-surface mutations.
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md),
[STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

The long-form cell and drawing reference is split so scalar cell writes, link/comment flows, and
drawing authoring can evolve without collapsing into one monolith.

| Detailed reference | Owns |
|:-------------------|:-----|
| [CELL_VALUE_MUTATIONS.md](./CELL_VALUE_MUTATIONS.md) | `SET_CELL`, `SET_RANGE`, `SET_ARRAY_FORMULA`, `CLEAR_ARRAY_FORMULA`, and `CLEAR_RANGE` |
| [LINK_AND_COMMENT_MUTATIONS.md](./LINK_AND_COMMENT_MUTATIONS.md) | hyperlink and classic-comment authoring and clearing |
| [DRAWING_MUTATIONS.md](./DRAWING_MUTATIONS.md) | pictures, shapes, embedded objects, charts, signature lines, anchor moves, and drawing deletes |
