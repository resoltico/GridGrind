---
afad: "3.3"
version: "1.0.0"
domain: OPERATIONS
updated: "2026-03-24"
route:
  keywords: [gridgrind, operations, set-cell, set-range, apply-style, ensure-sheet, append-row, clear-range, evaluate-formulas, auto-size-columns, request, json, protocol]
  questions: ["what operations does gridgrind support", "how do I set a cell value", "how do I apply a style", "how do I write a range", "what is the request format", "what fields does SET_RANGE accept"]
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
Create a new blank workbook.

```json
{ "mode": "EXISTING", "path": "path/to/workbook.xlsx" }
```
Open an existing `.xlsx` file.

---

## Persistence

```json
{ "mode": "SAVE_AS", "path": "path/to/output.xlsx" }
```
Write the workbook to the given path, creating parent directories as needed.

```json
{ "mode": "OVERWRITE", "path": "path/to/existing.xlsx" }
```
Overwrite an existing file at the given path.

Omit `persistence` entirely to run a read-only analysis without saving.

---

## Cell Values

Used in `SET_CELL`, `SET_RANGE`, and `APPEND_ROW`:

```json
{ "type": "TEXT",    "text": "Origin"  }
{ "type": "NUMBER",  "number": 8.40    }
{ "type": "BOOLEAN", "bool": true      }
{ "type": "FORMULA", "formula": "SUM(B2:B3)" }
{ "type": "BLANK"                      }
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
