---
afad: "3.4"
version: "0.5.0"
domain: OPERATIONS
updated: "2026-03-25"
route:
  keywords: [gridgrind, operations, set-cell, set-range, apply-style, ensure-sheet, rename-sheet, delete-sheet, move-sheet, merge-cells, unmerge-cells, set-column-width, set-row-height, freeze-panes, append-row, clear-range, evaluate-formulas, auto-size-columns, request, json, protocol]
  questions: ["what operations does gridgrind support", "how do I rename a sheet", "how do I delete a sheet", "how do I move a sheet", "how do I merge cells", "how do I set a column width", "how do I freeze panes", "how do I set a cell value", "how do I apply a style", "how do I write a range", "what is the request format", "what fields does SET_RANGE accept"]
---

# Operations Reference

**Purpose**: Complete reference for all GridGrind request fields and operation types.
**Prerequisites**: Familiarity with the [README](../README.md) quick start.

---

## Request Structure

```json
{
  "protocolVersion": "V1",
  "source":      { ... },
  "persistence": { ... },
  "operations":  [ ... ],
  "analysis":    { ... }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `protocolVersion` | No | Wire-contract version. Defaults to `V1`. Include it — future breaking revisions will be explicit. |
| `source` | Yes | Where the workbook comes from. |
| `persistence` | No | Where and whether to save. Omit to run operations without saving. |
| `operations` | No | Ordered list of workbook mutations. |
| `analysis` | No | What to inspect after operations complete. |

---

## Source

```json
{ "mode": "NEW" }
```
Create a new blank `.xlsx` workbook.

```json
{ "mode": "EXISTING", "path": "path/to/workbook.xlsx" }
```
Open an existing `.xlsx` file.

GridGrind supports `.xlsx` only. Paths ending in `.xls`, `.xlsm`, `.xlsb`, or any other
non-`.xlsx` extension are rejected as invalid requests.

---

## Persistence

```json
{ "mode": "SAVE_AS", "path": "path/to/output.xlsx" }
```
Write the workbook to the given path, creating parent directories as needed.

The save path must end in `.xlsx`.

```json
{ "mode": "OVERWRITE" }
```
Overwrite the source file (requires `source.mode=EXISTING`).

Omit `persistence` entirely to run a read-only analysis without saving.

---

## Cell Values

Used in `SET_CELL`, `SET_RANGE`, and `APPEND_ROW`:

```json
{ "type": "TEXT",      "text": "Origin"              }
{ "type": "NUMBER",    "number": 8.40                }
{ "type": "BOOLEAN",   "bool": true                  }
{ "type": "FORMULA",   "formula": "SUM(B2:B3)"       }
{ "type": "DATE",      "date": "2026-03-25"           }
{ "type": "DATE_TIME", "dateTime": "2026-03-25T10:15:30" }
{ "type": "BLANK"                                     }
```

---

## Operations

### ENSURE_SHEET

Create a sheet if it does not already exist. Does nothing if the sheet exists.

```json
{ "type": "ENSURE_SHEET", "sheetName": "Inventory" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Name of the sheet to create. |

---

### RENAME_SHEET

Rename an existing sheet. The source sheet must already exist. `newSheetName` must be a valid
Excel sheet name and must not conflict with another sheet name.

```json
{
  "type": "RENAME_SHEET",
  "sheetName": "Archive",
  "newSheetName": "History"
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing sheet to rename. |
| `newSheetName` | Yes | New sheet name. Must be valid and unique. |

---

### DELETE_SHEET

Delete an existing sheet.

```json
{ "type": "DELETE_SHEET", "sheetName": "Scratch" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing sheet to delete. |

---

### MOVE_SHEET

Move an existing sheet to a zero-based workbook position. `targetIndex` is evaluated at the time
the operation runs, after all earlier operations in the same request.

```json
{
  "type": "MOVE_SHEET",
  "sheetName": "History",
  "targetIndex": 0
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing sheet to move. |
| `targetIndex` | Yes | Zero-based destination index. Must be between `0` and `sheetCount - 1`. |

---

### MERGE_CELLS

Merge a rectangular A1-style range into one displayed region. The range must span at least two
cells. Repeating the same merge is a no-op, but overlapping a different merged region fails.

```json
{ "type": "MERGE_CELLS", "sheetName": "Inventory", "range": "A1:C1" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `range` | Yes | A1-notation rectangular range to merge. Must span at least two cells. |

---

### UNMERGE_CELLS

Remove the merged region whose coordinates exactly match `range`. GridGrind does not partially
unmerge intersecting regions; the range must identify the merged region exactly.

```json
{ "type": "UNMERGE_CELLS", "sheetName": "Inventory", "range": "A1:C1" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `range` | Yes | Exact A1-notation merged range to remove. |

---

### SET_COLUMN_WIDTH

Set the width of one or more contiguous columns. Column indexes are zero-based and inclusive.
`widthCharacters` is expressed in Excel character units and converted to Apache POI raw width
units with `round(widthCharacters * 256)`.

```json
{
  "type": "SET_COLUMN_WIDTH",
  "sheetName": "Inventory",
  "firstColumnIndex": 0,
  "lastColumnIndex": 2,
  "widthCharacters": 16.0
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `firstColumnIndex` | Yes | Zero-based first column index. |
| `lastColumnIndex` | Yes | Zero-based last column index. Must be greater than or equal to `firstColumnIndex`. |
| `widthCharacters` | Yes | Positive Excel character width. Must be finite and at most `255.0`. |

---

### SET_ROW_HEIGHT

Set the height of one or more contiguous rows. Row indexes are zero-based and inclusive.
`heightPoints` uses Excel point units and is passed to Apache POI via `setHeightInPoints`.

```json
{
  "type": "SET_ROW_HEIGHT",
  "sheetName": "Inventory",
  "firstRowIndex": 0,
  "lastRowIndex": 3,
  "heightPoints": 28.5
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `firstRowIndex` | Yes | Zero-based first row index. |
| `lastRowIndex` | Yes | Zero-based last row index. Must be greater than or equal to `firstRowIndex`. |
| `heightPoints` | Yes | Positive Excel row height in points. Must be finite. |

---

### FREEZE_PANES

Freeze panes using explicit split and visible-origin coordinates. `splitColumn` and `splitRow`
define the frozen boundary. `leftmostColumn` and `topRow` define the first visible column and row
in the scrollable pane after the split.

```json
{
  "type": "FREEZE_PANES",
  "sheetName": "Inventory",
  "splitColumn": 1,
  "splitRow": 1,
  "leftmostColumn": 1,
  "topRow": 1
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `splitColumn` | Yes | Zero-based frozen column count. |
| `splitRow` | Yes | Zero-based frozen row count. |
| `leftmostColumn` | Yes | Zero-based first visible column in the scrollable pane. |
| `topRow` | Yes | Zero-based first visible row in the scrollable pane. |

Constraints:

- `splitColumn` and `splitRow` must not both be `0`.
- If `splitColumn` is `0`, `leftmostColumn` must also be `0`.
- If `splitRow` is `0`, `topRow` must also be `0`.
- When `splitColumn > 0`, `leftmostColumn` must be greater than or equal to `splitColumn`.
- When `splitRow > 0`, `topRow` must be greater than or equal to `splitRow`.

---

### SET_CELL

Write a typed value to a single cell using A1 notation.

```json
{
  "type": "SET_CELL",
  "sheetName": "Inventory",
  "address": "B4",
  "value": { "type": "FORMULA", "formula": "SUM(B2:B3)" }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `address` | Yes | A1-notation cell address. |
| `value` | Yes | Typed cell value. |

---

### SET_RANGE

Write a rectangular grid of typed values. The `rows` matrix must exactly match the dimensions of
`range`. Row count must equal `range` row span; column count of each row must equal `range` column span.

```json
{
  "type": "SET_RANGE",
  "sheetName": "Inventory",
  "range": "A1:C3",
  "rows": [
    [{ "type": "TEXT", "text": "Origin" }, { "type": "TEXT", "text": "Kilos" }, { "type": "TEXT", "text": "Cost/kg" }],
    [{ "type": "TEXT", "text": "Ethiopia Yirgacheffe" }, { "type": "NUMBER", "number": 150 }, { "type": "NUMBER", "number": 8.40 }],
    [{ "type": "TEXT", "text": "Colombia Huila" }, { "type": "NUMBER", "number": 200 }, { "type": "NUMBER", "number": 7.80 }]
  ]
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `range` | Yes | A1-notation rectangular range (e.g. `A1:C3`). |
| `rows` | Yes | 2-D array of typed cell values. Dimensions must match `range`. |

---

### CLEAR_RANGE

Remove all values and styles from a rectangular range.

```json
{ "type": "CLEAR_RANGE", "sheetName": "Inventory", "range": "A2:C10" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `range` | Yes | Range to clear. |

---

### APPLY_STYLE

Apply a style patch to a rectangular range. Only fields present in `style` are changed;
unspecified style properties are left untouched.

```json
{
  "type": "APPLY_STYLE",
  "sheetName": "Inventory",
  "range": "A1:C1",
  "style": {
    "bold": true,
    "horizontalAlignment": "CENTER",
    "verticalAlignment": "CENTER",
    "wrapText": true
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `range` | Yes | Range to style. |
| `style` | Yes | Partial style patch. All fields optional. |

Style fields:

| Field | Type | Values |
|:------|:-----|:-------|
| `numberFormat` | string | Any Excel number format string, e.g. `"#,##0.00"`, `"yyyy-mm-dd"` |
| `bold` | boolean | `true` / `false` |
| `italic` | boolean | `true` / `false` |
| `wrapText` | boolean | `true` / `false` |
| `horizontalAlignment` | string | `"LEFT"`, `"CENTER"`, `"RIGHT"`, `"GENERAL"` |
| `verticalAlignment` | string | `"TOP"`, `"CENTER"`, `"BOTTOM"` |

---

### APPEND_ROW

Append a single row of typed values after the last populated row in a sheet.

```json
{
  "type": "APPEND_ROW",
  "sheetName": "Inventory",
  "values": [
    { "type": "TEXT",   "text": "Guatemala Antigua" },
    { "type": "NUMBER", "number": 100 },
    { "type": "NUMBER", "number": 9.20 }
  ]
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `values` | Yes | Row of typed cell values. |

---

### AUTO_SIZE_COLUMNS

Resize columns to fit their content. Applies to all columns with data on the sheet.

```json
{ "type": "AUTO_SIZE_COLUMNS", "sheetName": "Inventory" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |

Note: `AUTO_SIZE_COLUMNS` and `SET_COLUMN_WIDTH` can be combined in the same request. Since
operations run in order, whichever appears later wins.

---

### EVALUATE_FORMULAS

Force recalculation of all formulas in the workbook before save. Use this when formulas depend on
data written earlier in the same request.

```json
{ "type": "EVALUATE_FORMULAS" }
```

No additional fields.

---

### FORCE_FORMULA_RECALCULATION_ON_OPEN

Mark the workbook so Excel recalculates all formulas when the file is opened. Use this as an
alternative to `EVALUATE_FORMULAS` when server-side recalculation is not required.

```json
{ "type": "FORCE_FORMULA_RECALCULATION_ON_OPEN" }
```

No additional fields.

---

## Analysis

```json
{
  "analysis": {
    "sheets": [
      {
        "sheetName": "Inventory",
        "cells": ["A1", "B4"],
        "previewRowCount": 5,
        "previewColumnCount": 3
      }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheets` | Yes | List of sheet analysis requests. |
| `sheetName` | Yes | Sheet to analyze. |
| `cells` | No | Specific A1-notation cells to include in the response with full detail. |
| `previewRowCount` | No | Number of rows to include in the sheet preview. |
| `previewColumnCount` | No | Number of columns to include in the sheet preview. |

The preview includes styled blank cells so template-like workbooks remain visible to agents.
