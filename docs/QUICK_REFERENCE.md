---
afad: "3.4"
version: "0.7.0"
domain: QUICK_REFERENCE
updated: "2026-03-26"
route:
  keywords: [gridgrind, quick-reference, snippets, json, operations, copy-paste, ensure-sheet, rename-sheet, delete-sheet, move-sheet, merge-cells, unmerge-cells, set-column-width, set-row-height, freeze-panes, set-cell, set-range, set-hyperlink, clear-hyperlink, set-comment, clear-comment, set-named-range, delete-named-range, apply-style, append-row, clear-range, evaluate-formulas]
  questions: ["gridgrind json snippets", "how do I write a cell in gridgrind", "gridgrind copy paste examples", "gridgrind hyperlink example", "gridgrind comment example", "gridgrind named range example"]
---

# Quick Reference

Copy-paste JSON fragments for every GridGrind operation. All snippets are valid
standalone operation objects for use inside the `operations` array.

GridGrind supports `.xlsx` workbooks only. Use `.xlsx` paths for `source.path` and
`persistence.path`; `.xls`, `.xlsm`, and `.xlsb` are rejected.

---

## Request Skeleton

```json
{
  "protocolVersion": "V1",
  "source": { "mode": "NEW" },
  "persistence": { "mode": "SAVE_AS", "path": "output.xlsx" },
  "operations": [],
  "analysis": { "sheets": [] }
}
```

---

## Source

```json
{ "mode": "NEW" }
{ "mode": "EXISTING", "path": "path/to/file.xlsx" }
```

## Persistence

```json
{ "mode": "SAVE_AS",   "path": "path/to/output.xlsx" }
{ "mode": "OVERWRITE" }
```

---

## ENSURE_SHEET

```json
{ "type": "ENSURE_SHEET", "sheetName": "Sheet1" }
```

## RENAME_SHEET

```json
{ "type": "RENAME_SHEET", "sheetName": "Archive", "newSheetName": "History" }
```

## DELETE_SHEET

```json
{ "type": "DELETE_SHEET", "sheetName": "Scratch" }
```

## MOVE_SHEET

```json
{ "type": "MOVE_SHEET", "sheetName": "History", "targetIndex": 0 }
```

## MERGE_CELLS

```json
{ "type": "MERGE_CELLS", "sheetName": "Sheet1", "range": "A1:C1" }
```

## UNMERGE_CELLS

```json
{ "type": "UNMERGE_CELLS", "sheetName": "Sheet1", "range": "A1:C1" }
```

## SET_COLUMN_WIDTH

```json
{
  "type": "SET_COLUMN_WIDTH",
  "sheetName": "Sheet1",
  "firstColumnIndex": 0,
  "lastColumnIndex": 2,
  "widthCharacters": 16.0
}
```

## SET_ROW_HEIGHT

```json
{
  "type": "SET_ROW_HEIGHT",
  "sheetName": "Sheet1",
  "firstRowIndex": 0,
  "lastRowIndex": 3,
  "heightPoints": 28.5
}
```

## FREEZE_PANES

```json
{
  "type": "FREEZE_PANES",
  "sheetName": "Sheet1",
  "splitColumn": 1,
  "splitRow": 1,
  "leftmostColumn": 1,
  "topRow": 1
}
```

## SET_CELL

```json
{ "type": "SET_CELL", "sheetName": "Sheet1", "address": "A1", "value": { "type": "TEXT",      "text": "Hello"                        } }
{ "type": "SET_CELL", "sheetName": "Sheet1", "address": "B1", "value": { "type": "NUMBER",    "number": 42.0                         } }
{ "type": "SET_CELL", "sheetName": "Sheet1", "address": "C1", "value": { "type": "BOOLEAN",   "bool": true                           } }
{ "type": "SET_CELL", "sheetName": "Sheet1", "address": "D1", "value": { "type": "FORMULA",   "formula": "SUM(B1:B10)"               } }
{ "type": "SET_CELL", "sheetName": "Sheet1", "address": "E1", "value": { "type": "DATE",      "date": "2026-03-25"                   } }
{ "type": "SET_CELL", "sheetName": "Sheet1", "address": "F1", "value": { "type": "DATE_TIME", "dateTime": "2026-03-25T10:15:30"      } }
{ "type": "SET_CELL", "sheetName": "Sheet1", "address": "G1", "value": { "type": "BLANK"                                             } }
```

## SET_RANGE

```json
{
  "type": "SET_RANGE",
  "sheetName": "Sheet1",
  "range": "A1:C2",
  "rows": [
    [{ "type": "TEXT", "text": "A" }, { "type": "TEXT", "text": "B" }, { "type": "TEXT", "text": "C" }],
    [{ "type": "NUMBER", "number": 1 }, { "type": "NUMBER", "number": 2 }, { "type": "NUMBER", "number": 3 }]
  ]
}
```

## CLEAR_RANGE

```json
{ "type": "CLEAR_RANGE", "sheetName": "Sheet1", "range": "A1:Z100" }
```

`CLEAR_RANGE` removes value, style, hyperlink, and comment state from every cell in the range.

## SET_HYPERLINK

```json
{ "type": "SET_HYPERLINK", "sheetName": "Sheet1", "address": "A1", "target": { "type": "URL", "target": "https://example.com/report" } }
{ "type": "SET_HYPERLINK", "sheetName": "Sheet1", "address": "A2", "target": { "type": "EMAIL", "email": "team@example.com" } }
{ "type": "SET_HYPERLINK", "sheetName": "Sheet1", "address": "A3", "target": { "type": "FILE", "target": "reports/monthly.xlsx" } }
{ "type": "SET_HYPERLINK", "sheetName": "Sheet1", "address": "A4", "target": { "type": "DOCUMENT", "target": "Summary!B4" } }
```

## CLEAR_HYPERLINK

```json
{ "type": "CLEAR_HYPERLINK", "sheetName": "Sheet1", "address": "A1" }
```

## SET_COMMENT

```json
{
  "type": "SET_COMMENT",
  "sheetName": "Sheet1",
  "address": "B2",
  "comment": {
    "text": "Reviewed by finance.",
    "author": "GridGrind",
    "visible": true
  }
}
```

If `visible` is omitted, it defaults to `false`.

## CLEAR_COMMENT

```json
{ "type": "CLEAR_COMMENT", "sheetName": "Sheet1", "address": "B2" }
```

## APPLY_STYLE

```json
{
  "type": "APPLY_STYLE",
  "sheetName": "Sheet1",
  "range": "A1:C1",
  "style": {
    "bold": true,
    "italic": false,
    "wrapText": true,
    "numberFormat": "#,##0.00",
    "fontName": "Aptos",
    "fontHeight": { "type": "POINTS", "points": 13 },
    "fontColor": "#1F4E78",
    "underline": true,
    "strikeout": false,
    "fillColor": "#FFF2CC",
    "horizontalAlignment": "CENTER",
    "verticalAlignment": "CENTER",
    "border": {
      "all": { "style": "THIN" },
      "right": { "style": "DOUBLE" }
    }
  }
}
```

`horizontalAlignment` values: `"LEFT"` `"CENTER"` `"RIGHT"` `"GENERAL"`
`verticalAlignment` values:   `"TOP"` `"CENTER"` `"BOTTOM"`
`fontColor` and `fillColor` use `#RRGGBB`.
`fontHeight` accepts either `{ "type": "POINTS", "points": 11.5 }` or `{ "type": "TWIPS", "twips": 230 }`.
`border.all` sets the default side style; explicit sides override it.

## APPEND_ROW

```json
{
  "type": "APPEND_ROW",
  "sheetName": "Sheet1",
  "values": [
    { "type": "TEXT", "text": "Row label" },
    { "type": "NUMBER", "number": 100 }
  ]
}
```

## AUTO_SIZE_COLUMNS

```json
{ "type": "AUTO_SIZE_COLUMNS", "sheetName": "Sheet1" }
```

## SET_NAMED_RANGE

```json
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetTotal",
  "scope": { "type": "WORKBOOK" },
  "target": { "sheetName": "Budget", "range": "B4" }
}
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetTable",
  "scope": { "type": "SHEET", "sheetName": "Budget" },
  "target": { "sheetName": "Budget", "range": "A1:B4" }
}
```

Named-range targets are normalized to top-left:`bottom-right` order. For example, `"B4:A1"` is
accepted and stored as `"A1:B4"`.

## DELETE_NAMED_RANGE

```json
{
  "type": "DELETE_NAMED_RANGE",
  "name": "BudgetTotal",
  "scope": { "type": "WORKBOOK" }
}
```

`SET_COLUMN_WIDTH.widthCharacters` is converted to POI width units with `round(widthCharacters * 256)`.
`SET_ROW_HEIGHT.heightPoints` uses Excel point units. `UNMERGE_CELLS` requires an exact merged-range match.

## EVALUATE_FORMULAS

```json
{ "type": "EVALUATE_FORMULAS" }
```

## FORCE_FORMULA_RECALCULATION_ON_OPEN

```json
{ "type": "FORCE_FORMULA_RECALCULATION_ON_OPEN" }
```

---

## Analysis

```json
{
  "sheets": [
    {
      "sheetName": "Sheet1",
      "cells": ["A1", "B4", "C10"],
      "previewRowCount": 5,
      "previewColumnCount": 4
    }
  ],
  "namedRanges": { "mode": "ALL" }
}
```

Named-range analysis snippets:

```json
{ "mode": "NONE" }
{ "mode": "ALL" }
{
  "mode": "SELECTED",
  "selectors": [
    { "type": "WORKBOOK_SCOPE", "name": "BudgetTotal" },
    { "type": "SHEET_SCOPE", "sheetName": "Budget", "name": "BudgetTable" }
  ]
}
```
