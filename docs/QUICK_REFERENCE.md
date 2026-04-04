---
afad: "3.4"
version: "0.26.0"
domain: QUICK_REFERENCE
updated: "2026-04-02"
route:
  keywords: [gridgrind, quick-reference, snippets, json, operations, reads, introspection, analysis, copy-paste, ensure-sheet, rename-sheet, delete-sheet, move-sheet, copy-sheet, set-active-sheet, set-selected-sheets, set-sheet-visibility, set-sheet-protection, clear-sheet-protection, merge-cells, unmerge-cells, set-column-width, set-row-height, set-sheet-pane, set-sheet-zoom, set-print-layout, clear-print-layout, freeze-panes, split-panes, set-cell, set-range, set-hyperlink, clear-hyperlink, set-comment, clear-comment, set-data-validation, clear-data-validations, set-autofilter, clear-autofilter, set-table, delete-table, set-named-range, delete-named-range, apply-style, append-row, clear-range, evaluate-formulas, get-cells, get-window, get-print-layout, get-data-validations, get-autofilters, get-tables, get-sheet-schema, analyze-autofilter-health, analyze-table-health, analyze-workbook-findings]
  questions: ["gridgrind json snippets", "how do I write a cell in gridgrind", "gridgrind copy paste examples", "gridgrind copy sheet example", "gridgrind active sheet example", "gridgrind selected sheets example", "gridgrind sheet visibility example", "gridgrind sheet protection example", "gridgrind hyperlink example", "gridgrind comment example", "gridgrind table example", "gridgrind autofilter example", "gridgrind named range example", "what do gridgrind reads look like"]
---

# Quick Reference

Copy-paste JSON fragments for every GridGrind operation. All snippets are valid
standalone operation objects for use inside the `operations` array.

GridGrind supports `.xlsx` workbooks only. Use `.xlsx` paths for `source.path` and
`persistence.path`; `.xls`, `.xlsm`, and `.xlsb` are rejected.

The artifact can emit the current contract directly:

```bash
gridgrind --print-request-template
gridgrind --print-protocol-catalog
```

`--print-protocol-catalog` returns machine-readable field descriptors. Every catalog entry states
which fields are required or optional and, for polymorphic fields, which nested or plain type
group supplies the accepted JSON shape.

---

## Request Skeleton

```json
{
  "protocolVersion": "V1",
  "source": { "type": "NEW" },
  "persistence": { "type": "SAVE_AS", "path": "output.xlsx" },
  "operations": [],
  "reads": []
}
```

Path model:
- `source.path` opens an existing workbook from that path.
- `SAVE_AS.path` writes a new workbook to that path and creates missing parent directories.
- `OVERWRITE` writes back to `source.path`; it does not accept its own `path` field.
- Relative paths in `--request`, `--response`, `source.path`, and `persistence.path` resolve from
  the current working directory.

---

## Source

```json
{ "type": "NEW" }
{ "type": "EXISTING", "path": "path/to/file.xlsx" }
```

Relative `path` values resolve from the current working directory.

## Persistence

```json
{ "type": "SAVE_AS",   "path": "path/to/output.xlsx" }
{ "type": "OVERWRITE" }
```

`SAVE_AS.path` writes a new file. `OVERWRITE` writes back to `source.path` and does not accept a
separate `path` field. `SAVE_AS` creates missing parent directories automatically.

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

`DELETE_SHEET` rejects deleting the last remaining sheet or the last visible sheet in a workbook.

## MOVE_SHEET

```json
{ "type": "MOVE_SHEET", "sheetName": "History", "targetIndex": 0 }
```

## COPY_SHEET

```json
{
  "type": "COPY_SHEET",
  "sourceSheetName": "Budget",
  "newSheetName": "Budget Review",
  "position": { "type": "AT_INDEX", "targetIndex": 1 }
}
```

## SET_ACTIVE_SHEET

```json
{ "type": "SET_ACTIVE_SHEET", "sheetName": "Budget Review" }
```

## SET_SELECTED_SHEETS

```json
{ "type": "SET_SELECTED_SHEETS", "sheetNames": ["Budget", "Budget Review"] }
```

## SET_SHEET_VISIBILITY

```json
{ "type": "SET_SHEET_VISIBILITY", "sheetName": "Archive", "visibility": "VERY_HIDDEN" }
```

## SET_SHEET_PROTECTION

```json
{
  "type": "SET_SHEET_PROTECTION",
  "sheetName": "Budget Review",
  "protection": {
    "autoFilterLocked": true,
    "deleteColumnsLocked": true,
    "deleteRowsLocked": true,
    "formatCellsLocked": true,
    "formatColumnsLocked": false,
    "formatRowsLocked": false,
    "insertColumnsLocked": true,
    "insertHyperlinksLocked": true,
    "insertRowsLocked": true,
    "objectsLocked": true,
    "pivotTablesLocked": true,
    "scenariosLocked": true,
    "selectLockedCellsLocked": true,
    "selectUnlockedCellsLocked": false,
    "sortLocked": true
  }
}
```

## CLEAR_SHEET_PROTECTION

```json
{ "type": "CLEAR_SHEET_PROTECTION", "sheetName": "Budget Review" }
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

## SET_SHEET_PANE

```json
{
  "type": "SET_SHEET_PANE",
  "sheetName": "Sheet1",
  "pane": {
    "type": "FROZEN",
    "splitColumn": 1,
    "splitRow": 1,
    "leftmostColumn": 1,
    "topRow": 1
  }
}
```

```json
{
  "type": "SET_SHEET_PANE",
  "sheetName": "Sheet1",
  "pane": {
    "type": "SPLIT",
    "xSplitPosition": 2400,
    "ySplitPosition": 1800,
    "leftmostColumn": 2,
    "topRow": 3,
    "activePane": "LOWER_RIGHT"
  }
}
```

## SET_SHEET_ZOOM

```json
{ "type": "SET_SHEET_ZOOM", "sheetName": "Sheet1", "zoomPercent": 125 }
```

## SET_PRINT_LAYOUT

```json
{
  "type": "SET_PRINT_LAYOUT",
  "sheetName": "Sheet1",
  "printLayout": {
    "printArea": { "type": "RANGE", "range": "A1:F40" },
    "orientation": "LANDSCAPE",
    "scaling": { "type": "FIT", "widthPages": 1, "heightPages": 0 },
    "repeatingRows": { "type": "BAND", "firstRowIndex": 0, "lastRowIndex": 1 },
    "repeatingColumns": { "type": "BAND", "firstColumnIndex": 0, "lastColumnIndex": 0 },
    "header": { "left": "Inventory", "center": "Q2", "right": "Internal" },
    "footer": { "left": "", "center": "Prepared by GridGrind", "right": "Page 1" }
  }
}
```

## CLEAR_PRINT_LAYOUT

```json
{ "type": "CLEAR_PRINT_LAYOUT", "sheetName": "Sheet1" }
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

A leading `=` in `FORMULA` values is accepted and stripped automatically. `"=SUM(B1:B10)"` and `"SUM(B1:B10)"` are equivalent.
Existing style, hyperlink, and comment state on the targeted cell is preserved. `DATE` and
`DATE_TIME` writes keep existing presentation state and only layer the required number format on
top.

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
Cells that have never been written are silently skipped.

## SET_HYPERLINK

```json
{ "type": "SET_HYPERLINK", "sheetName": "Sheet1", "address": "A1", "target": { "type": "URL", "target": "https://example.com/report" } }
{ "type": "SET_HYPERLINK", "sheetName": "Sheet1", "address": "A2", "target": { "type": "EMAIL", "email": "team@example.com" } }
{ "type": "SET_HYPERLINK", "sheetName": "Sheet1", "address": "A3", "target": { "type": "FILE", "path": "reports/monthly.xlsx" } }
{ "type": "SET_HYPERLINK", "sheetName": "Sheet1", "address": "A4", "target": { "type": "DOCUMENT", "target": "Summary!B4" } }
```

`FILE.path` accepts either a plain path or a `file:` URI. Reads return a normalized plain path
string in `path`.

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

## SET_DATA_VALIDATION

```json
{
  "type": "SET_DATA_VALIDATION",
  "sheetName": "Sheet1",
  "range": "B2:B20",
  "validation": {
    "rule": {
      "type": "EXPLICIT_LIST",
      "values": ["Queued", "Done", "Blocked"]
    },
    "allowBlank": true,
    "prompt": {
      "title": "Status",
      "text": "Pick one workflow state."
    },
    "errorAlert": {
      "style": "STOP",
      "title": "Invalid status",
      "text": "Use one of the allowed values."
    }
  }
}
{
  "type": "SET_DATA_VALIDATION",
  "sheetName": "Sheet1",
  "range": "C2:C20",
  "validation": {
    "rule": {
      "type": "WHOLE_NUMBER",
      "operator": "BETWEEN",
      "formula1": "1",
      "formula2": "5"
    }
  }
}
```

Overlapping existing rules are normalized so the written rule becomes authoritative on the target
range.

## CLEAR_DATA_VALIDATIONS

```json
{
  "type": "CLEAR_DATA_VALIDATIONS",
  "sheetName": "Sheet1",
  "selection": { "type": "ALL" }
}
{
  "type": "CLEAR_DATA_VALIDATIONS",
  "sheetName": "Sheet1",
  "selection": {
    "type": "SELECTED",
    "ranges": ["B4", "C8:D10"]
  }
}
```

## SET_CONDITIONAL_FORMATTING

```json
{
  "type": "SET_CONDITIONAL_FORMATTING",
  "sheetName": "Sheet1",
  "conditionalFormatting": {
    "ranges": ["D2:D20"],
    "rules": [
      {
        "type": "CELL_VALUE_RULE",
        "operator": "GREATER_THAN",
        "formula1": "8",
        "stopIfTrue": false,
        "style": {
          "fillColor": "#FDE9D9",
          "fontColor": "#9C0006",
          "bold": true
        }
      }
    ]
  }
}
{
  "type": "SET_CONDITIONAL_FORMATTING",
  "sheetName": "Sheet1",
  "conditionalFormatting": {
    "ranges": ["A2:D20"],
    "rules": [
      {
        "type": "FORMULA_RULE",
        "formula": "$D2=\"Blocked\"",
        "stopIfTrue": true,
        "style": {
          "fillColor": "#FFF2CC",
          "fontColor": "#7F6000",
          "italic": true
        }
      }
    ]
  }
}
```

Write support covers `FORMULA_RULE` and `CELL_VALUE_RULE`.

## CLEAR_CONDITIONAL_FORMATTING

```json
{
  "type": "CLEAR_CONDITIONAL_FORMATTING",
  "sheetName": "Sheet1",
  "selection": { "type": "ALL" }
}
{
  "type": "CLEAR_CONDITIONAL_FORMATTING",
  "sheetName": "Sheet1",
  "selection": {
    "type": "SELECTED",
    "ranges": ["D2:D20", "F2:F20"]
  }
}
```

## SET_AUTOFILTER

```json
{
  "type": "SET_AUTOFILTER",
  "sheetName": "Sheet1",
  "range": "A1:C20"
}
```

The range must include a nonblank header row and must not overlap any existing table range.

## CLEAR_AUTOFILTER

```json
{ "type": "CLEAR_AUTOFILTER", "sheetName": "Sheet1" }
```

This clears only the sheet-level autofilter. Table-owned autofilters remain attached to their
tables.

## SET_TABLE

```json
{
  "type": "SET_TABLE",
  "table": {
    "name": "DispatchQueue",
    "sheetName": "Sheet1",
    "range": "A1:C20",
    "showTotalsRow": false,
    "style": { "type": "NONE" }
  }
}
{
  "type": "SET_TABLE",
  "table": {
    "name": "DispatchQueue",
    "sheetName": "Sheet1",
    "range": "A1:C20",
    "showTotalsRow": true,
    "style": {
      "type": "NAMED",
      "name": "TableStyleMedium2",
      "showFirstColumn": false,
      "showLastColumn": false,
      "showRowStripes": true,
      "showColumnStripes": false
    }
  }
}
```

Table names are workbook-global. Header cells in the first row of the range must be nonblank and
unique case-insensitively. Overlapping different-name tables are rejected. Later value writes and
style patches that touch the table header row keep the table's persisted column metadata aligned
with the visible header cells.

## DELETE_TABLE

```json
{
  "type": "DELETE_TABLE",
  "name": "DispatchQueue",
  "sheetName": "Sheet1"
}
```

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

Metadata-only blank rows do not affect the append position.
If the append lands on cells that already exist only because of style, hyperlink, or comment
state, that existing state is preserved.

## AUTO_SIZE_COLUMNS

```json
{ "type": "AUTO_SIZE_COLUMNS", "sheetName": "Sheet1" }
```

Sizing is deterministic and content-based, so headless and Docker runs match local runs.

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

`SET_COLUMN_WIDTH.widthCharacters` is converted to POI width units with `round(widthCharacters * 256)`. Must be > 0 and ≤ 255.0.
`SET_ROW_HEIGHT.heightPoints` uses Excel point units. Must be > 0 and ≤ 1,638.35 (32,767 twips). `UNMERGE_CELLS` requires an exact merged-range match.
`DELETE_SHEET` returns `INVALID_REQUEST` when the sheet to delete is the only remaining sheet or the last visible sheet in the workbook.

## EVALUATE_FORMULAS

```json
{ "type": "EVALUATE_FORMULAS" }
```

## FORCE_FORMULA_RECALCULATION_ON_OPEN

```json
{ "type": "FORCE_FORMULA_RECALCULATION_ON_OPEN" }
```

---

## Reads

Each read is a standalone object for use inside the top-level `reads` array. Every read must have
a caller-defined `requestId`.

## GET_WORKBOOK_SUMMARY

```json
{ "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" }
```

## GET_NAMED_RANGES

```json
{
  "type": "GET_NAMED_RANGES",
  "requestId": "named-ranges",
  "selection": { "type": "ALL" }
}
{
  "type": "GET_NAMED_RANGES",
  "requestId": "selected-named-ranges",
  "selection": {
    "type": "SELECTED",
    "selectors": [
      { "type": "WORKBOOK_SCOPE", "name": "BudgetTotal" },
      { "type": "SHEET_SCOPE", "sheetName": "Budget", "name": "BudgetTable" }
    ]
  }
}
```

## GET_SHEET_SUMMARY

```json
{ "type": "GET_SHEET_SUMMARY", "requestId": "sheet-summary", "sheetName": "Sheet1" }
```

## GET_CELLS

```json
{
  "type": "GET_CELLS",
  "requestId": "cells",
  "sheetName": "Sheet1",
  "addresses": ["A1", "B4", "C10"]
}
```

## GET_WINDOW

```json
{
  "type": "GET_WINDOW",
  "requestId": "window",
  "sheetName": "Sheet1",
  "topLeftAddress": "A1",
  "rowCount": 5,
  "columnCount": 4
}
```

`rowCount * columnCount` must not exceed 250,000.

## GET_MERGED_REGIONS

```json
{ "type": "GET_MERGED_REGIONS", "requestId": "merged", "sheetName": "Sheet1" }
```

## GET_HYPERLINKS

Response hyperlinks reuse the same discriminated shape as `SET_HYPERLINK` targets. `FILE`
responses use the `path` field with a normalized plain path string.

```json
{
  "type": "GET_HYPERLINKS",
  "requestId": "hyperlinks",
  "sheetName": "Sheet1",
  "selection": { "type": "ALL_USED_CELLS" }
}
{
  "type": "GET_HYPERLINKS",
  "requestId": "selected-hyperlinks",
  "sheetName": "Sheet1",
  "selection": {
    "type": "SELECTED",
    "addresses": ["A1", "B4"]
  }
}
```

## GET_COMMENTS

```json
{
  "type": "GET_COMMENTS",
  "requestId": "comments",
  "sheetName": "Sheet1",
  "selection": { "type": "ALL_USED_CELLS" }
}
{
  "type": "GET_COMMENTS",
  "requestId": "selected-comments",
  "sheetName": "Sheet1",
  "selection": {
    "type": "SELECTED",
    "addresses": ["A1", "B4"]
  }
}
```

## GET_SHEET_LAYOUT

```json
{ "type": "GET_SHEET_LAYOUT", "requestId": "layout", "sheetName": "Sheet1" }
```

## GET_PRINT_LAYOUT

```json
{ "type": "GET_PRINT_LAYOUT", "requestId": "print-layout", "sheetName": "Sheet1" }
```

## GET_DATA_VALIDATIONS

```json
{
  "type": "GET_DATA_VALIDATIONS",
  "requestId": "data-validations",
  "sheetName": "Sheet1",
  "selection": { "type": "ALL" }
}
{
  "type": "GET_DATA_VALIDATIONS",
  "requestId": "selected-data-validations",
  "sheetName": "Sheet1",
  "selection": {
    "type": "SELECTED",
    "ranges": ["B2:B20", "C2:C20"]
  }
}
```

## GET_AUTOFILTERS

```json
{
  "type": "GET_AUTOFILTERS",
  "requestId": "autofilters",
  "sheetName": "Sheet1"
}
```

Entries are `SHEET` or `TABLE`.

## GET_CONDITIONAL_FORMATTING

```json
{
  "type": "GET_CONDITIONAL_FORMATTING",
  "requestId": "conditional-formatting",
  "sheetName": "Sheet1",
  "selection": { "type": "ALL" }
}
{
  "type": "GET_CONDITIONAL_FORMATTING",
  "requestId": "selected-conditional-formatting",
  "sheetName": "Sheet1",
  "selection": {
    "type": "SELECTED",
    "ranges": ["A2:D20", "F2:F20"]
  }
}
```

Read entries may report `FORMULA_RULE`, `CELL_VALUE_RULE`, `COLOR_SCALE_RULE`, `DATA_BAR_RULE`,
`ICON_SET_RULE`, or `UNSUPPORTED_RULE`.

## GET_TABLES

```json
{
  "type": "GET_TABLES",
  "requestId": "tables",
  "selection": { "type": "ALL" }
}
{
  "type": "GET_TABLES",
  "requestId": "selected-tables",
  "selection": {
    "type": "BY_NAMES",
    "names": ["DispatchQueue", "Trips"]
  }
}
```

## GET_FORMULA_SURFACE

```json
{
  "type": "GET_FORMULA_SURFACE",
  "requestId": "formula-surface",
  "selection": {
    "type": "ALL"
  }
}
{
  "type": "GET_FORMULA_SURFACE",
  "requestId": "selected-formula-surface",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Sheet1", "Sheet2"]
  }
}
```

## GET_SHEET_SCHEMA

```json
{
  "type": "GET_SHEET_SCHEMA",
  "requestId": "schema",
  "sheetName": "Sheet1",
  "topLeftAddress": "A1",
  "rowCount": 5,
  "columnCount": 4
}
```

`rowCount * columnCount` must not exceed 250,000. When the first row is entirely blank,
`dataRowCount` in the response is `0`.

## GET_NAMED_RANGE_SURFACE

```json
{
  "type": "GET_NAMED_RANGE_SURFACE",
  "requestId": "named-range-surface",
  "selection": { "type": "ALL" }
}
{
  "type": "GET_NAMED_RANGE_SURFACE",
  "requestId": "selected-named-range-surface",
  "selection": {
    "type": "SELECTED",
    "selectors": [
      { "type": "WORKBOOK_SCOPE", "name": "BudgetTotal" },
      { "type": "SHEET_SCOPE", "sheetName": "Budget", "name": "BudgetTable" }
    ]
  }
}
```

## ANALYZE_FORMULA_HEALTH

```json
{
  "type": "ANALYZE_FORMULA_HEALTH",
  "requestId": "formula-health",
  "selection": { "type": "ALL" }
}
```

## ANALYZE_DATA_VALIDATION_HEALTH

```json
{
  "type": "ANALYZE_DATA_VALIDATION_HEALTH",
  "requestId": "data-validation-health",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Sheet1", "Sheet2"]
  }
}
```

## ANALYZE_CONDITIONAL_FORMATTING_HEALTH

```json
{
  "type": "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
  "requestId": "conditional-formatting-health",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Sheet1", "Sheet2"]
  }
}
```

## ANALYZE_AUTOFILTER_HEALTH

```json
{
  "type": "ANALYZE_AUTOFILTER_HEALTH",
  "requestId": "autofilter-health",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Sheet1", "Sheet2"]
  }
}
```

## ANALYZE_TABLE_HEALTH

```json
{
  "type": "ANALYZE_TABLE_HEALTH",
  "requestId": "table-health",
  "selection": { "type": "ALL" }
}
{
  "type": "ANALYZE_TABLE_HEALTH",
  "requestId": "selected-table-health",
  "selection": {
    "type": "BY_NAMES",
    "names": ["DispatchQueue", "Trips"]
  }
}
```

## ANALYZE_HYPERLINK_HEALTH

```json
{
  "type": "ANALYZE_HYPERLINK_HEALTH",
  "requestId": "hyperlink-health",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Sheet1", "Sheet2"]
  }
}
```

Relative `FILE` targets are resolved against the workbook's persisted path when one exists.
Unsaved workbooks report unresolved relative file targets instead of treating them as healthy.

## ANALYZE_NAMED_RANGE_HEALTH

```json
{
  "type": "ANALYZE_NAMED_RANGE_HEALTH",
  "requestId": "named-range-health",
  "selection": { "type": "ALL" }
}
```

## ANALYZE_WORKBOOK_FINDINGS

```json
{
  "type": "ANALYZE_WORKBOOK_FINDINGS",
  "requestId": "workbook-findings"
}
```

`ANALYZE_WORKBOOK_FINDINGS` aggregates every shipped health family: formula, data validation,
conditional formatting, autofilter, table, hyperlink, and named range.

Selection snippets:

```json
{ "type": "ALL_USED_CELLS" }
{
  "type": "SELECTED",
  "addresses": ["A1", "B4", "C10"]
}
```

```json
{ "type": "ALL" }
{
  "type": "SELECTED",
  "sheetNames": ["Sheet1", "Sheet2"]
}
```

```json
{ "type": "ALL" }
{
  "type": "SELECTED",
  "selectors": [
    { "type": "WORKBOOK_SCOPE", "name": "BudgetTotal" },
    { "type": "SHEET_SCOPE", "sheetName": "Budget", "name": "BudgetTable" }
  ]
}
```

```json
{ "type": "ALL" }
{
  "type": "BY_NAMES",
  "names": ["DispatchQueue", "Trips"]
}
```
```
