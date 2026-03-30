---
afad: "3.4"
version: "0.14.0"
domain: OPERATIONS
updated: "2026-03-29"
route:
  keywords: [gridgrind, operations, reads, introspection, analysis, set-cell, set-range, apply-style, ensure-sheet, rename-sheet, delete-sheet, move-sheet, merge-cells, unmerge-cells, set-column-width, set-row-height, freeze-panes, append-row, clear-range, evaluate-formulas, auto-size-columns, get-cells, get-window, get-sheet-layout, get-sheet-schema, analyze-workbook-findings, request, json, protocol]
  questions: ["what operations does gridgrind support", "what reads does gridgrind support", "how do I rename a sheet", "how do I delete a sheet", "how do I move a sheet", "how do I merge cells", "how do I set a column width", "how do I freeze panes", "how do I set a cell value", "how do I apply a style", "how do I write a range", "what is the request format", "what fields does SET_RANGE accept", "what does GET_CELLS accept"]
---

# Operations Reference

**Purpose**: Complete reference for all GridGrind request fields and operation types.
**Prerequisites**: Familiarity with the [README](../README.md) quick start.
**Limits**: See [LIMITATIONS.md](LIMITATIONS.md) for all hard ceilings — window cell count, Excel maximums, and memory constraints.

The CLI can emit the current minimal request and the full machine-readable protocol inventory:

```bash
gridgrind --print-request-template
gridgrind --print-protocol-catalog
```

---

## Request Structure

```json
{
  "protocolVersion": "V1",
  "source":      { ... },
  "persistence": { ... },
  "operations":  [ ... ],
  "reads":       [ ... ]
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `protocolVersion` | No | Wire-contract version. Defaults to `V1`. Include it — future breaking revisions will be explicit. |
| `source` | Yes | Where the workbook comes from. |
| `persistence` | No | Where and whether to save. Omit to run operations without saving. |
| `operations` | No | Ordered list of workbook mutations. |
| `reads` | No | Ordered post-mutation introspection and analysis operations. |

Every tagged request union uses `type` as its discriminator field: `source`, `persistence`,
`operations`, `reads`, cell values, hyperlink targets, selections, and named-range scopes.

---

## Source

```json
{ "type": "NEW" }
```
Create a new blank `.xlsx` workbook. The workbook starts with zero sheets; use `ENSURE_SHEET` to
create the first sheet before writing any cells.

```json
{ "type": "EXISTING", "path": "path/to/workbook.xlsx" }
```
Open an existing `.xlsx` file.

GridGrind supports `.xlsx` only. Paths ending in `.xls`, `.xlsm`, `.xlsb`, or any other
non-`.xlsx` extension are rejected as invalid requests.

---

## Persistence

The response `persistence.type` field always echoes the request `persistence.type` value, making
it straightforward to correlate request and response: a `SAVE_AS` request yields a `SAVE_AS`
response, an `OVERWRITE` request yields an `OVERWRITE` response, and a `NONE` request yields a
`NONE` response.

```json
{ "type": "SAVE_AS", "path": "path/to/output.xlsx" }
```
Write the workbook to the given path, creating parent directories as needed.

The save path must end in `.xlsx`.

The response includes two path fields:
- `requestedPath` — the literal `path` string from the request.
- `executionPath` — the absolute normalized path where the file was actually written.

They are identical when an absolute path with no `..` segments is supplied. They differ when a
relative path (e.g. `"report.xlsx"`) or a path containing `..` segments is used.

```json
{ "type": "OVERWRITE" }
```
Overwrite the source file (requires `source.type=EXISTING`). The response includes `sourcePath`
(the original source path string) and `executionPath` (the absolute normalized path).

Omit `persistence` entirely or use `{ "type": "NONE" }` to run mutations and reads without
saving.

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

All mutation operations (`SET_CELL`, `SET_RANGE`, `CLEAR_RANGE`, `APPLY_STYLE`, `SET_HYPERLINK`,
`CLEAR_HYPERLINK`, `SET_COMMENT`, `CLEAR_COMMENT`, `APPEND_ROW`, `AUTO_SIZE_COLUMNS`) require the
target sheet to already exist. Use `ENSURE_SHEET` before the first write to any sheet.

```json
{ "type": "ENSURE_SHEET", "sheetName": "Inventory" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Name of the sheet to create. Maximum 31 characters. |

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

Remove all values, styles, hyperlinks, and comments from a rectangular range. Cells that have
never been written are silently skipped; this operation never fails because a cell does not
physically exist.

```json
{ "type": "CLEAR_RANGE", "sheetName": "Inventory", "range": "A2:C10" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `range` | Yes | Range to clear. |

---

### SET_HYPERLINK

Attach or replace a hyperlink on one cell. The cell is created if needed.

```json
{
  "type": "SET_HYPERLINK",
  "sheetName": "Inventory",
  "address": "A1",
  "target": {
    "type": "URL",
    "target": "https://example.com/catalog"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `address` | Yes | A1-notation cell address. |
| `target` | Yes | Typed hyperlink target payload. |

Supported target variants:

- `{"type":"URL","target":"https://example.com/report"}`
- `{"type":"EMAIL","email":"team@example.com"}`
- `{"type":"FILE","target":"shared/report.xlsx"}`
- `{"type":"DOCUMENT","target":"Inventory!B4"}`

---

### CLEAR_HYPERLINK

Remove any hyperlink attached to one cell. The cell value, style, and comment are left unchanged.
No-op when the cell does not physically exist.

```json
{ "type": "CLEAR_HYPERLINK", "sheetName": "Inventory", "address": "A1" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `address` | Yes | A1-notation cell address. |

---

### SET_COMMENT

Attach or replace one plain-text comment on one cell. The cell is created if needed.

```json
{
  "type": "SET_COMMENT",
  "sheetName": "Inventory",
  "address": "B4",
  "comment": {
    "text": "Verified by finance.",
    "author": "GridGrind",
    "visible": true
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `address` | Yes | A1-notation cell address. |
| `comment` | Yes | Plain-text comment payload. |

`comment.visible` defaults to `false` when omitted.

---

### CLEAR_COMMENT

Remove any comment attached to one cell. The cell value, style, and hyperlink are left unchanged.
No-op when the cell does not physically exist.

```json
{ "type": "CLEAR_COMMENT", "sheetName": "Inventory", "address": "B4" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `address` | Yes | A1-notation cell address. |

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
    "fontName": "Aptos",
    "fontHeight": { "type": "POINTS", "points": 13 },
    "fontColor": "#FFFFFF",
    "fillColor": "#1F4E78",
    "horizontalAlignment": "CENTER",
    "verticalAlignment": "CENTER",
    "wrapText": true,
    "border": {
      "all": { "style": "THIN" },
      "bottom": { "style": "DOUBLE" }
    }
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
| `fontName` | string | Excel font family name, e.g. `"Aptos"` |
| `fontHeight` | object | Typed font height. See below. |
| `fontColor` | string | RGB hex string in `#RRGGBB` form. Lowercase input is normalized to uppercase. |
| `underline` | boolean | `true` adds a single underline, `false` removes it |
| `strikeout` | boolean | `true` / `false` |
| `fillColor` | string | Solid foreground fill color in `#RRGGBB` form |
| `border` | object | Nested border patch. See below. |

Font-height input:

| Field | Type | Values |
|:------|:-----|:-------|
| `fontHeight.type` | string | `"POINTS"` or `"TWIPS"` |
| `fontHeight.points` | number | Required when `type` is `"POINTS"`. Must resolve exactly to whole twips, e.g. `11`, `11.5`, `13.25` |
| `fontHeight.twips` | integer | Required when `type` is `"TWIPS"`. Positive exact twip value where `20` twips = `1` point |

Color notes:

- `fontColor` and `fillColor` must match `#RRGGBB`.
- Lowercase hex input is accepted and normalized to uppercase.
- This contract does not support alpha channels, theme colors, indexed colors, or gradient fills.

Fill notes:

- `fillColor` always means a solid foreground fill.
- Non-solid patterns and background-fill authoring are not part of the public contract.
- Cell snapshot reads (GET_CELLS, GET_WINDOW, GET_SHEET_SCHEMA) report `style.fontHeight` as a
  plain object with both `twips` and `points` fields: `{"twips": 260, "points": 13}`. This is
  not the same as the write-side `FontHeightInput` discriminated format. See the Cell snapshot
  shape section under GET_CELLS for the full write-vs-read comparison.

Border patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `all` | object | Optional default for every border side not explicitly set in the same patch |
| `top` | object | Optional side-specific override |
| `right` | object | Optional side-specific override |
| `bottom` | object | Optional side-specific override |
| `left` | object | Optional side-specific override |

Each border-side object currently contains one field:

| Field | Type | Values |
|:------|:-----|:-------|
| `style` | string | `"NONE"`, `"THIN"`, `"MEDIUM"`, `"DASHED"`, `"DOTTED"`, `"THICK"`, `"DOUBLE"`, `"HAIR"`, `"MEDIUM_DASHED"`, `"DASH_DOT"`, `"MEDIUM_DASH_DOT"`, `"DASH_DOT_DOT"`, `"MEDIUM_DASH_DOT_DOT"`, `"SLANTED_DASH_DOT"` |

Border notes:

- `border` must set at least one of `all`, `top`, `right`, `bottom`, or `left`.
- `border.all` acts as the default style for every side not explicitly set in the same patch.
- Explicit side settings override `border.all`.
- Border colors and diagonal borders are not part of the public contract.

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

### SET_NAMED_RANGE

Create or replace one named range in workbook scope or sheet scope. Targets are explicit
sheet-qualified cells or rectangular ranges.

```json
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetTotal",
  "scope": { "type": "WORKBOOK" },
  "target": { "sheetName": "Budget", "range": "B4" }
}
```

```json
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetTable",
  "scope": { "type": "SHEET", "sheetName": "Budget" },
  "target": { "sheetName": "Budget", "range": "A1:B4" }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Defined-name identifier. Must not collide with A1 or R1C1 reference syntax and must not use the reserved `_xlnm.` prefix. |
| `scope` | Yes | Workbook or sheet scope payload. |
| `target` | Yes | Explicit sheet-qualified cell or rectangular range target. Reversed range endpoints are normalized to top-left:`bottom-right`. |

---

### DELETE_NAMED_RANGE

Delete one existing named range from workbook scope or sheet scope.

```json
{
  "type": "DELETE_NAMED_RANGE",
  "name": "BudgetTotal",
  "scope": { "type": "WORKBOOK" }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Defined-name identifier to delete. |
| `scope` | Yes | Workbook or sheet scope of the exact name to delete. |

---

## Reads

Reads are ordered, explicit post-mutation requests. Every read must include a caller-defined
`requestId`, and every result echoes that `requestId` back in the successful response.

Read categories:

- Introspection: exact workbook facts with no higher-level interpretation.
- Analysis: finding-bearing workbook conclusions built on top of introspection.

```json
{
  "reads": [
    { "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" },
    {
      "type": "GET_WINDOW",
      "requestId": "inventory-window",
      "sheetName": "Inventory",
      "topLeftAddress": "A1",
      "rowCount": 5,
      "columnCount": 3
    },
    {
      "type": "GET_SHEET_SCHEMA",
      "requestId": "inventory-schema",
      "sheetName": "Inventory",
      "topLeftAddress": "A1",
      "rowCount": 5,
      "columnCount": 3
    }
  ]
}
```

### GET_WORKBOOK_SUMMARY

Returns workbook-level summary facts such as sheet order, named-range count, and the workbook
force-recalculation flag.

```json
{ "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" }
```

### GET_NAMED_RANGES

Returns exact named-range reports selected by workbook-wide or exact-selector input.

```json
{
  "type": "GET_NAMED_RANGES",
  "requestId": "named-ranges",
  "selection": { "type": "ALL" }
}
```

```json
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

| Field | Required | Description |
|:------|:---------|:------------|
| `requestId` | Yes | Stable caller-defined correlation identifier. |
| `selection` | Yes | `ALL` or exact named-range selectors. |

Selected named-range reads fail with `NAMED_RANGE_NOT_FOUND` when any selector does not match an
existing named range.

### GET_SHEET_SUMMARY

Returns structural summary facts for one sheet.

```json
{
  "type": "GET_SHEET_SUMMARY",
  "requestId": "sheet-summary",
  "sheetName": "Inventory"
}
```

Response field semantics:

| Field | Description |
|:------|:------------|
| `physicalRowCount` | Number of physically materialized rows in the sheet. Rows are sparse; a sheet with data only in row 1 and row 100 has `physicalRowCount=2`. |
| `lastRowIndex` | Zero-based index of the last populated row. `-1` when the sheet is empty. Not the same as `physicalRowCount - 1` when rows are sparse. |
| `lastColumnIndex` | Zero-based index of the last populated column across all rows. `-1` when the sheet is empty. |

### GET_CELLS

Returns exact cell snapshots for one sheet and an ordered list of A1 addresses.

```json
{
  "type": "GET_CELLS",
  "requestId": "cells",
  "sheetName": "Inventory",
  "addresses": ["A1", "B4", "C10"]
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `requestId` | Yes | Stable caller-defined correlation identifier. |
| `sheetName` | Yes | Sheet that owns the requested cells. |
| `addresses` | Yes | Ordered non-empty list of unique A1 cell addresses. |

`GET_CELLS` returns a blank-typed snapshot for any valid address that has never been written.
An address that is not valid A1 notation (e.g. `BADADDR`, `A0`) or that exceeds the Excel 2007
sheet boundary (row > 1,048,575 or column > 16,383, e.g. `A1048577`, `XFE1`) returns
`INVALID_CELL_ADDRESS`, not a blank. It fails with `SHEET_NOT_FOUND` when the target sheet does
not exist.

#### Cell snapshot shape

Every cell in a `GET_CELLS`, `GET_WINDOW`, or `GET_SHEET_SCHEMA` response has the following
common fields:

| Field | Description |
|:------|:------------|
| `address` | A1-style cell address. |
| `declaredType` | Raw Excel cell type: `BLANK`, `STRING`, `NUMERIC`, `BOOLEAN`, `FORMULA`, or `ERROR`. |
| `effectiveType` | For non-formula cells: same as `declaredType`. For formula cells: always `FORMULA` — the evaluated result type is in `evaluation.effectiveType`. |
| `displayValue` | Formatted string as Excel would render it. |
| `style` | Style snapshot: `numberFormat`, `bold`, `italic`, `wrapText`, `horizontalAlignment`, `verticalAlignment`, `fontName`, `fontHeight`, `fontColor`, `underline`, `strikeout`, `fillColor`, border sides. |
| `metadata` | Hyperlink and comment metadata attached at snapshot time. |

Type-specific value fields (present only on matching `effectiveType`):

| Field | Present when | Description |
|:------|:-------------|:------------|
| `stringValue` | `effectiveType=STRING` | The string content. |
| `numberValue` | `effectiveType=NUMERIC` | The numeric value as a double. |
| `booleanValue` | `effectiveType=BOOLEAN` | `true` or `false`. |
| `errorValue` | `effectiveType=ERROR` | The error string (e.g. `#DIV/0!`). |
| `formula` | `declaredType=FORMULA` | The formula text without the leading `=`. |
| `evaluation` | `declaredType=FORMULA` | Nested cell snapshot of the evaluated result. |

**fontHeight read-back shape:** `style.fontHeight` is a plain object with both `twips` and
`points` fields: `{"twips": 260, "points": 13}`. This differs from the write-side
`FontHeightInput` format which uses a discriminated `{"type": "POINTS", "points": 13}` object.
Agents that read a font height and want to write it back must use the `FontHeightInput` write
format, not the read-back shape.

### GET_WINDOW

Returns a rectangular top-left-anchored window of cell snapshots. The window includes styled blank
cells so template-like workbooks remain visible.

`rowCount * columnCount` must not exceed 250,000. The window must not extend beyond the Excel
2007 sheet boundary (rows 0–1,048,575, columns 0–16,383); requests that overflow are rejected
with `INVALID_REQUEST`.

```json
{
  "type": "GET_WINDOW",
  "requestId": "window",
  "sheetName": "Inventory",
  "topLeftAddress": "A1",
  "rowCount": 5,
  "columnCount": 3
}
```

Response shape: `{ "window": { "sheetName": "...", "rows": [ { "cells": [...] } ] } }`. The
top-level key is `window` and cells are nested under `window.rows[N].cells`. This differs from
`GET_CELLS` where cells are directly under the top-level `cells` key.

### GET_MERGED_REGIONS

Returns the exact merged regions defined on one sheet.

```json
{
  "type": "GET_MERGED_REGIONS",
  "requestId": "merged",
  "sheetName": "Inventory"
}
```

### GET_HYPERLINKS

Returns hyperlink metadata for selected cells on one sheet.

```json
{
  "type": "GET_HYPERLINKS",
  "requestId": "hyperlinks",
  "sheetName": "Inventory",
  "selection": { "type": "ALL_USED_CELLS" }
}
```

```json
{
  "type": "GET_HYPERLINKS",
  "requestId": "selected-hyperlinks",
  "sheetName": "Inventory",
  "selection": {
    "type": "SELECTED",
    "addresses": ["A1", "B4"]
  }
}
```

### GET_COMMENTS

Returns comment metadata for selected cells on one sheet.

```json
{
  "type": "GET_COMMENTS",
  "requestId": "comments",
  "sheetName": "Inventory",
  "selection": { "type": "ALL_USED_CELLS" }
}
```

```json
{
  "type": "GET_COMMENTS",
  "requestId": "selected-comments",
  "sheetName": "Inventory",
  "selection": {
    "type": "SELECTED",
    "addresses": ["A1", "B4"]
  }
}
```

Cell-selection payloads use:

```json
{ "type": "ALL_USED_CELLS" }
{
  "type": "SELECTED",
  "addresses": ["A1", "B4"]
}
```

### GET_SHEET_LAYOUT

Returns freeze-pane state plus explicit row-height and column-width facts for one sheet.

```json
{
  "type": "GET_SHEET_LAYOUT",
  "requestId": "layout",
  "sheetName": "Inventory"
}
```

### GET_FORMULA_SURFACE

Groups formula usage across one or more sheets.

```json
{
  "type": "GET_FORMULA_SURFACE",
  "requestId": "formula-surface",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "GET_FORMULA_SURFACE",
  "requestId": "selected-formula-surface",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Inventory", "Summary"]
  }
}
```

Sheet-selection payloads use:

```json
{ "type": "ALL" }
{
  "type": "SELECTED",
  "sheetNames": ["Inventory", "Summary"]
}
```

### GET_SHEET_SCHEMA

Infers a simple schema from one rectangular sheet window. The first row of the window is treated
as a header row; `dataRowCount` in the response is `rowCount - 1`. When the header row is
entirely blank, `dataRowCount` is `0`.

`rowCount * columnCount` must not exceed 250,000. Requests that exceed this limit are rejected
with `INVALID_REQUEST`.

`dominantType` per column is `null` when all data cells are blank, or when two or more types tie
for the highest count. Formula cells contribute their evaluated result type (e.g. `NUMERIC`,
`STRING`) to `observedTypes` and `dominantType`, not `FORMULA`.

```json
{
  "type": "GET_SHEET_SCHEMA",
  "requestId": "schema",
  "sheetName": "Inventory",
  "topLeftAddress": "A1",
  "rowCount": 5,
  "columnCount": 3
}
```

### GET_NAMED_RANGE_SURFACE

Summarizes the scope and backing kind of selected named ranges.

```json
{
  "type": "GET_NAMED_RANGE_SURFACE",
  "requestId": "named-range-surface",
  "selection": { "type": "ALL" }
}
```

```json
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

### ANALYZE_FORMULA_HEALTH

Reports finding-bearing formula health across one or more sheets. This is where volatile
functions, formula-error results, and evaluation failures surface.

```json
{
  "type": "ANALYZE_FORMULA_HEALTH",
  "requestId": "formula-health",
  "selection": { "type": "ALL" }
}
```

### ANALYZE_HYPERLINK_HEALTH

Reports hyperlink findings such as malformed external targets or broken document destinations.

```json
{
  "type": "ANALYZE_HYPERLINK_HEALTH",
  "requestId": "hyperlink-health",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Inventory", "Summary"]
  }
}
```

### ANALYZE_NAMED_RANGE_HEALTH

Reports named-range findings such as broken references, unresolved targets, and scope shadowing.

```json
{
  "type": "ANALYZE_NAMED_RANGE_HEALTH",
  "requestId": "named-range-health",
  "selection": { "type": "ALL" }
}
```

### ANALYZE_WORKBOOK_FINDINGS

Runs the first finding-bearing analysis family across the workbook and returns one aggregated
finding list.

```json
{
  "type": "ANALYZE_WORKBOOK_FINDINGS",
  "requestId": "workbook-findings"
}
```
