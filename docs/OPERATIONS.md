---
afad: "3.5"
version: "0.41.0"
domain: OPERATIONS
updated: "2026-04-13"
route:
  keywords: [gridgrind, operations, reads, introspection, analysis, set-cell, set-range, apply-style, ensure-sheet, rename-sheet, delete-sheet, move-sheet, copy-sheet, set-active-sheet, set-selected-sheets, set-sheet-visibility, set-sheet-protection, clear-sheet-protection, set-workbook-protection, clear-workbook-protection, merge-cells, unmerge-cells, set-column-width, set-row-height, set-sheet-pane, set-sheet-zoom, set-print-layout, clear-print-layout, freeze-panes, split-panes, set-data-validation, set-autofilter, clear-autofilter, set-table, delete-table, set-pivot-table, delete-pivot-table, set-picture, set-chart, set-shape, set-embedded-object, set-drawing-object-anchor, delete-drawing-object, append-row, clear-range, evaluate-formulas, auto-size-columns, get-cells, get-window, get-print-layout, get-workbook-protection, get-data-validations, get-autofilters, get-tables, get-pivot-tables, get-drawing-objects, get-charts, get-drawing-object-payload, get-sheet-layout, get-sheet-schema, analyze-autofilter-health, analyze-table-health, analyze-pivot-table-health, analyze-workbook-findings, request, json, protocol, coordinates, rowindex, columnindex, warnings]
  questions: ["what operations does gridgrind support", "what reads does gridgrind support", "how do I rename a sheet", "how do I delete a sheet", "how do I move a sheet", "how do I copy a sheet", "how do I set the active sheet", "how do I set selected sheets", "how do I set sheet visibility", "how do I set sheet protection", "how do I set workbook protection", "how do I merge cells", "how do I set a column width", "how do I freeze panes", "how do I set split panes", "how do I set sheet zoom", "how do I set print layout", "how do I set a cell value", "how do I apply a style", "how do I write a range", "how do I create an autofilter in gridgrind", "how do I create a table in gridgrind", "how do I create a pivot table in gridgrind", "how do I read pivot tables in gridgrind", "how do I add a picture in gridgrind", "how do I author a chart in gridgrind", "how do I read charts in gridgrind", "how do I read drawing objects in gridgrind", "what is the request format", "what fields does SET_RANGE accept", "what does GET_CELLS accept", "which fields use A1 notation versus zero-based indexes", "how do I run workbook findings without saving"]
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

The protocol catalog is designed for black-box consumers. Each published type entry includes field
descriptors with required/optional status and recursive shape metadata, so fields like
`selection`, `target`, `value`, `style`, and `scope` point directly at the nested/plain type
group that defines their accepted JSON structure.

---

## Request Structure

```json
{
  "protocolVersion": "V1",
  "source":      { ... },
  "persistence": { ... },
  "executionMode": { ... },
  "formulaEnvironment": { ... },
  "operations":  [ ... ],
  "reads":       [ ... ]
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `protocolVersion` | No | Wire-contract version. Defaults to `V1`. Include it — future breaking revisions will be explicit. |
| `source` | Yes | Where the workbook comes from. |
| `persistence` | No | Where and whether to save. Omit to run operations without saving. |
| `executionMode` | No | Optional low-memory read and write mode selection. Omit for the default full-XSSF path. |
| `formulaEnvironment` | No | Request-scoped evaluator configuration for external workbook bindings, missing-workbook policy, and template-backed UDF toolpacks. |
| `operations` | No | Ordered list of workbook mutations. |
| `reads` | No | Ordered post-mutation introspection and analysis operations. |

Every tagged request union uses `type` as its discriminator field: `source`, `persistence`,
`operations`, `reads`, cell values, hyperlink targets, selections, and named-range scopes.

### Formula Environment

`formulaEnvironment` is optional. Omit it for the default evaluator. Supply it when server-side
formula evaluation needs external workbook bindings, cached-value fallback for unresolved external
references, or template-backed UDFs.

```json
{
  "formulaEnvironment": {
    "externalWorkbooks": [
      {
        "workbookName": "rates.xlsx",
        "path": "fixtures/rates.xlsx"
      }
    ],
    "missingWorkbookPolicy": "USE_CACHED_VALUE",
    "udfToolpacks": [
      {
        "name": "math",
        "functions": [
          {
            "name": "DOUBLE",
            "minimumArgumentCount": 1,
            "formulaTemplate": "ARG1*2"
          }
        ]
      }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `externalWorkbooks` | No | Workbook-name to path bindings used to satisfy formulas such as `[rates.xlsx]Sheet1!A1`. |
| `missingWorkbookPolicy` | No | `ERROR` or `USE_CACHED_VALUE`. Defaults to `ERROR`. |
| `udfToolpacks` | No | Named collections of template-backed UDFs. |

For `udfToolpacks.functions`, `maximumArgumentCount` is optional and defaults to
`minimumArgumentCount`. `formulaTemplate` may reference `ARG1`, `ARG2`, and higher placeholders.

### Execution Mode

`executionMode` is optional. Omit it for the default `FULL_XSSF` request path.

```json
{
  "executionMode": {
    "readMode": "EVENT_READ",
    "writeMode": "STREAMING_WRITE"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `readMode` | No | `FULL_XSSF` or `EVENT_READ`. Defaults to `FULL_XSSF`. |
| `writeMode` | No | `FULL_XSSF` or `STREAMING_WRITE`. Defaults to `FULL_XSSF`. |

- `readMode: EVENT_READ` selects the low-memory XSSF event-model reader. It supports only
  `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY` (`LIM-019`).
- `writeMode: STREAMING_WRITE` selects the low-memory SXSSF writer. It requires `source.type:
  NEW`, supports only `ENSURE_SHEET`, `APPEND_ROW`, and `FORCE_FORMULA_RECALC_ON_OPEN`, and
  requires at least one `ENSURE_SHEET` or `APPEND_ROW` (`LIM-020`).
- `EVENT_READ` can run directly against an existing workbook when the request is read-only and
  unsaved. If the request also performs full-XSSF mutations, GridGrind materializes the mutated
  workbook state and then performs the summary reads through the event model.
- `STREAMING_WRITE` can pair with either `readMode: FULL_XSSF` for broader readback or
  `readMode: EVENT_READ` for summary-only low-memory readback from the materialized streaming
  result.

## Coordinate Systems

GridGrind uses two coordinate conventions:

| Field pattern | Convention |
|:--------------|:-----------|
| `address` | A1 cell address, e.g. `B3` |
| `range` | A1 rectangular range, e.g. `A1:C4` |
| `*RowIndex` | Zero-based row index, e.g. `0 = Excel row 1` |
| `*ColumnIndex` | Zero-based column index, e.g. `0 = Excel column A` |

`first...` and `last...` index pairs are inclusive zero-based bands. Validation messages echo the
Excel-native equivalent inline, for example `firstRowIndex 5 (Excel row 6)` or
`firstColumnIndex 5 (Excel column F)`.

## Checking Workbook Health

Use `ANALYZE_WORKBOOK_FINDINGS` as the primary workbook-health check. Pair it with
`persistence.type=NONE` when you only need findings and do not want a saved workbook:

```json
{
  "source": { "type": "NEW" },
  "persistence": { "type": "NONE" },
  "operations": [],
  "reads": [
    { "type": "ANALYZE_WORKBOOK_FINDINGS", "requestId": "lint" }
  ]
}
```

Successful responses may include a `warnings` array. The current request-phase warning flags
same-request sheet names with spaces referenced in formulas without single quotes. Use
`'Sheet Name'!A1` syntax for those references.

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

Relative `path` values resolve from the current working directory.

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

Relative `path` values resolve from the current working directory.

The save path must end in `.xlsx`.

The response includes two path fields:
- `requestedPath` — the literal `path` string from the request.
- `executionPath` — the absolute normalized path where the file was actually written.

They are identical when an absolute path with no `..` segments is supplied. They differ when a
relative path (e.g. `"report.xlsx"`) or a path containing `..` segments is used.

```json
{ "type": "OVERWRITE" }
```
Overwrite the source file (requires `source.type=EXISTING`). `OVERWRITE` does not accept its own
`path` field; it always writes back to `source.path`. The response includes `sourcePath` (the
original source path string) and `executionPath` (the absolute normalized path).

Omit `persistence` entirely or use `{ "type": "NONE" }` to run mutations and reads without
saving.

---

## Cell Values

Used in `SET_CELL`, `SET_RANGE`, and `APPEND_ROW`:

```json
{ "type": "TEXT",      "text": "Origin"              }
{
  "type": "RICH_TEXT",
  "runs": [
    { "text": "Q2 ", "font": { "fontName": "Aptos", "fontColor": "#44546A" } },
    { "text": "Budget", "font": { "bold": true, "fontColor": "#C00000" } }
  ]
}
{ "type": "NUMBER",    "number": 8.40                }
{ "type": "BOOLEAN",   "bool": true                  }
{ "type": "FORMULA",   "formula": "SUM(B2:B3)"       }  // leading = is accepted and stripped
{ "type": "DATE",      "date": "2026-03-25"           }
{ "type": "DATE_TIME", "dateTime": "2026-03-25T10:15:30" }
{ "type": "BLANK"                                     }
```

`RICH_TEXT` writes an ordered, non-empty `runs` list. Every run must have non-empty `text`, and
the optional `font` object reuses the same font-field vocabulary as the nested style contract:
`bold`, `italic`, `fontName`, `fontHeight`, `fontColor`, `underline`, and `strikeout`.

---

## Operations

### ENSURE_SHEET

Create a sheet if it does not already exist. Does nothing if the sheet exists.

All mutation operations (`SET_CELL`, `SET_RANGE`, `CLEAR_RANGE`, `APPLY_STYLE`,
`SET_DATA_VALIDATION`, `CLEAR_DATA_VALIDATIONS`, `SET_CONDITIONAL_FORMATTING`,
`CLEAR_CONDITIONAL_FORMATTING`, `SET_HYPERLINK`, `CLEAR_HYPERLINK`, `SET_COMMENT`,
`CLEAR_COMMENT`, `SET_PICTURE`, `SET_CHART`, `SET_SHAPE`, `SET_EMBEDDED_OBJECT`,
`SET_DRAWING_OBJECT_ANCHOR`, `DELETE_DRAWING_OBJECT`, `SET_AUTOFILTER`,
`CLEAR_AUTOFILTER`, `APPEND_ROW`, `AUTO_SIZE_COLUMNS`) require the target sheet to already exist.
`SET_TABLE` and `SET_PIVOT_TABLE` also require the target `table.sheetName` or
`pivotTable.sheetName` to exist. Use `ENSURE_SHEET` before the first write to any sheet.

```json
{ "type": "ENSURE_SHEET", "sheetName": "Inventory" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Name of the sheet to create. Must be 1 to 31 characters and must not contain `:` `\` `/` `?` `*` `[` `]`, or begin or end with a single quote. |

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

Delete an existing sheet. A workbook must retain at least one sheet and at least one visible
sheet, so attempting to delete the last remaining sheet or the last visible sheet returns
`INVALID_REQUEST`.

```json
{ "type": "DELETE_SHEET", "sheetName": "Scratch" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing sheet to delete. Must not be the only remaining sheet or the last visible sheet. |

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

### COPY_SHEET

Copy one sheet into a new visible, unselected sheet. The copied sheet is placed either at the end
of workbook order or at an explicit zero-based index. GridGrind preserves supported sheet-local
content such as formulas, validations, conditional formatting, comments, hyperlinks, merged
regions, tables, sheet-scoped names, protection metadata, and layout state. Drawing-family
content such as pictures and charts remains outside the current copy contract, so sheet copy is
complete for non-drawing workbook-core structures and intentionally partial for drawing families.

```json
{
  "type": "COPY_SHEET",
  "sourceSheetName": "Budget",
  "newSheetName": "Budget Review",
  "position": { "type": "AT_INDEX", "targetIndex": 1 }
}
```

```json
{
  "type": "COPY_SHEET",
  "sourceSheetName": "Budget",
  "newSheetName": "Budget Snapshot"
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sourceSheetName` | Yes | Existing sheet to copy. |
| `newSheetName` | Yes | New unique destination sheet name. |
| `position` | No | Copy position. Omit to append at end. |

`position` variants:

- `{"type":"APPEND_AT_END"}`
- `{"type":"AT_INDEX","targetIndex":1}`

---

### SET_ACTIVE_SHEET

Set the active sheet. Hidden sheets cannot be activated. The active sheet is always selected.

```json
{ "type": "SET_ACTIVE_SHEET", "sheetName": "Budget Review" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing visible sheet to activate. |

---

### SET_SELECTED_SHEETS

Set the selected visible sheet set. Duplicate or unknown sheet names are rejected. Workbook
summary reads return `selectedSheetNames` in workbook order while preserving the chosen active
sheet as the primary selected tab.

```json
{ "type": "SET_SELECTED_SHEETS", "sheetNames": ["Budget", "Budget Review"] }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetNames` | Yes | Non-empty distinct list of existing visible sheets. |

---

### SET_SHEET_VISIBILITY

Set one sheet visibility state. A workbook must retain at least one visible sheet, so hiding the
last visible sheet is rejected.

```json
{ "type": "SET_SHEET_VISIBILITY", "sheetName": "Archive", "visibility": "VERY_HIDDEN" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `visibility` | Yes | `VISIBLE`, `HIDDEN`, or `VERY_HIDDEN`. |

---

### SET_SHEET_PROTECTION

Enable sheet protection with the exact supported lock flags and an optional password. When
`password` is provided, GridGrind writes a password hash into the sheet-protection XML while
preserving the explicit authored lock flags.

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
  },
  "password": "Sheet-2026"
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `protection` | Yes | Supported lock-flag payload. |
| `password` | No | Optional nonblank sheet-protection password. Omit it to protect the sheet without a stored password hash. |

---

### CLEAR_SHEET_PROTECTION

Disable sheet protection entirely and clear any stored password hash.

```json
{ "type": "CLEAR_SHEET_PROTECTION", "sheetName": "Budget Review" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |

---

### SET_WORKBOOK_PROTECTION

Enable workbook-level protection with authoritative lock flags and optional workbook or revisions
passwords. Omitted booleans normalize to `false`; omitted passwords clear the corresponding stored
hash.

```json
{
  "type": "SET_WORKBOOK_PROTECTION",
  "protection": {
    "structureLocked": true,
    "windowsLocked": false,
    "revisionsLocked": true,
    "workbookPassword": "Vault-2026",
    "revisionsPassword": "Revisions-2026"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `protection` | Yes | Workbook protection payload carrying structure/windows/revisions lock flags plus optional passwords. |

`protection` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `structureLocked` | No | Lock workbook structure. Defaults to `false`. |
| `windowsLocked` | No | Lock workbook windows. Defaults to `false`. |
| `revisionsLocked` | No | Lock workbook revisions. Defaults to `false`. |
| `workbookPassword` | No | Optional nonblank workbook-protection password. Omit to clear the workbook password hash. |
| `revisionsPassword` | No | Optional nonblank revisions-protection password. Omit to clear the revisions password hash. |

---

### CLEAR_WORKBOOK_PROTECTION

Disable workbook-level protection entirely and remove any stored workbook or revisions password
hashes.

```json
{ "type": "CLEAR_WORKBOOK_PROTECTION" }
```

No additional fields.

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
| `widthCharacters` | Yes | Positive Excel character width. Must be finite and > 0 and ≤ 255.0 (Excel column width limit). |

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
| `heightPoints` | Yes | Positive Excel row height in points. Must be finite and > 0 and ≤ 1,638.35 (Excel storage limit: 32,767 twips). |

---

### INSERT_ROWS

Insert one or more blank rows before `rowIndex`. `rowIndex` is zero-based and may point at the
current append position (`last existing row + 1`) but not beyond it.

```json
{
  "type": "INSERT_ROWS",
  "sheetName": "Inventory",
  "rowIndex": 2,
  "rowCount": 3
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `rowIndex` | Yes | Zero-based insertion point. Must be within the current row bounds or exactly one past the last existing row. |
| `rowCount` | Yes | Positive number of blank rows to insert. |

GridGrind rejects the edit when Apache POI would leave a moved table, sheet-owned autofilter, or
data validation stale (`LIM-016`).

---

### DELETE_ROWS

Delete one inclusive zero-based row band. Rows beneath the deleted band shift upward.

```json
{
  "type": "DELETE_ROWS",
  "sheetName": "Inventory",
  "rows": { "type": "BAND", "firstRowIndex": 4, "lastRowIndex": 6 }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `rows` | Yes | One inclusive row-band payload. |

GridGrind rejects the edit when it would move or truncate a table, sheet-owned autofilter, or
data validation (`LIM-016`), and it also rejects deletes that would truncate a range-backed named
range (`LIM-018`).

---

### SHIFT_ROWS

Move one inclusive zero-based row band by a signed non-zero `delta`. Positive values move rows
down; negative values move rows up.

```json
{
  "type": "SHIFT_ROWS",
  "sheetName": "Inventory",
  "rows": { "type": "BAND", "firstRowIndex": 1, "lastRowIndex": 3 },
  "delta": 2
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `rows` | Yes | One inclusive row-band payload. |
| `delta` | Yes | Signed non-zero row offset. |

GridGrind rejects the edit when it would move a table, sheet-owned autofilter, or data
validation (`LIM-016`), and it also rejects shifts that would partially move or overwrite a
range-backed named range outside the moved band (`LIM-018`).

---

### INSERT_COLUMNS

Insert one or more blank columns before `columnIndex`. `columnIndex` is zero-based and may point
at the current append position (`last existing column + 1`) but not beyond it.

```json
{
  "type": "INSERT_COLUMNS",
  "sheetName": "Inventory",
  "columnIndex": 1,
  "columnCount": 2
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `columnIndex` | Yes | Zero-based insertion point. Must be within the current column bounds or exactly one past the last existing column. |
| `columnCount` | Yes | Positive number of blank columns to insert. |

GridGrind rejects the edit when Apache POI would leave a moved table, sheet-owned autofilter, or
data validation stale (`LIM-016`), and it also rejects the edit when the workbook contains any
formula cells or formula-defined named ranges because Apache POI leaves some column references
stale after column moves (`LIM-017`).

---

### DELETE_COLUMNS

Delete one inclusive zero-based column band. Columns to the right of the deleted band shift left.

```json
{
  "type": "DELETE_COLUMNS",
  "sheetName": "Inventory",
  "columns": { "type": "BAND", "firstColumnIndex": 3, "lastColumnIndex": 4 }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `columns` | Yes | One inclusive column-band payload. |

GridGrind enforces ownership rejection (`LIM-016`), the formula-bearing workbook guard for column
structural edits (`LIM-017`), and a named-range guard that rejects deletes which would truncate a
range-backed named range (`LIM-018`).

---

### SHIFT_COLUMNS

Move one inclusive zero-based column band by a signed non-zero `delta`. Positive values move
columns right; negative values move columns left.

```json
{
  "type": "SHIFT_COLUMNS",
  "sheetName": "Inventory",
  "columns": { "type": "BAND", "firstColumnIndex": 0, "lastColumnIndex": 1 },
  "delta": -1
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `columns` | Yes | One inclusive column-band payload. |
| `delta` | Yes | Signed non-zero column offset. |

GridGrind enforces ownership rejection (`LIM-016`), the formula-bearing workbook guard for column
structural edits (`LIM-017`), and a named-range guard that rejects shifts which would partially
move or overwrite a range-backed named range outside the moved band (`LIM-018`).

---

### SET_ROW_VISIBILITY

Hide or unhide one inclusive zero-based row band.

```json
{
  "type": "SET_ROW_VISIBILITY",
  "sheetName": "Inventory",
  "rows": { "type": "BAND", "firstRowIndex": 5, "lastRowIndex": 7 },
  "hidden": true
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `rows` | Yes | One inclusive row-band payload. |
| `hidden` | Yes | `true` hides the rows; `false` unhides them. |

---

### SET_COLUMN_VISIBILITY

Hide or unhide one inclusive zero-based column band.

```json
{
  "type": "SET_COLUMN_VISIBILITY",
  "sheetName": "Inventory",
  "columns": { "type": "BAND", "firstColumnIndex": 2, "lastColumnIndex": 3 },
  "hidden": false
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `columns` | Yes | One inclusive column-band payload. |
| `hidden` | Yes | `true` hides the columns; `false` unhides them. |

---

### GROUP_ROWS

Apply one outline group to one inclusive zero-based row band. When `collapsed` is omitted it
defaults to `false`.

```json
{
  "type": "GROUP_ROWS",
  "sheetName": "Inventory",
  "rows": { "type": "BAND", "firstRowIndex": 8, "lastRowIndex": 10 },
  "collapsed": true
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `rows` | Yes | One inclusive row-band payload. |
| `collapsed` | No | When `true`, the newly grouped rows are collapsed immediately. Defaults to `false`. |

---

### UNGROUP_ROWS

Remove one outline group from one inclusive zero-based row band. GridGrind expands the band first
so hidden rows are not left stranded in a collapsed state.

```json
{
  "type": "UNGROUP_ROWS",
  "sheetName": "Inventory",
  "rows": { "type": "BAND", "firstRowIndex": 8, "lastRowIndex": 10 }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `rows` | Yes | One inclusive row-band payload. |

---

### GROUP_COLUMNS

Apply one outline group to one inclusive zero-based column band. When `collapsed` is omitted it
defaults to `false`.

```json
{
  "type": "GROUP_COLUMNS",
  "sheetName": "Inventory",
  "columns": { "type": "BAND", "firstColumnIndex": 4, "lastColumnIndex": 6 },
  "collapsed": true
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `columns` | Yes | One inclusive column-band payload. |
| `collapsed` | No | When `true`, the newly grouped columns are collapsed immediately. Defaults to `false`. |

---

### UNGROUP_COLUMNS

Remove one outline group from one inclusive zero-based column band. GridGrind expands the band
first so hidden columns are not left stranded in a collapsed state.

```json
{
  "type": "UNGROUP_COLUMNS",
  "sheetName": "Inventory",
  "columns": { "type": "BAND", "firstColumnIndex": 4, "lastColumnIndex": 6 }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `columns` | Yes | One inclusive column-band payload. |

---

### SET_SHEET_PANE

Apply one explicit pane state to a sheet. Use `NONE` to clear panes, `FROZEN` for freeze-pane
behavior, or `SPLIT` for Excel split panes.

```json
{
  "type": "SET_SHEET_PANE",
  "sheetName": "Inventory",
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
  "sheetName": "Inventory",
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

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `pane` | Yes | Pane payload with `type` `NONE`, `FROZEN`, or `SPLIT`. |

`pane.type = "NONE"` has no nested fields.

`pane.type = "FROZEN"` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `splitColumn` | Yes | Zero-based frozen column count. |
| `splitRow` | Yes | Zero-based frozen row count. |
| `leftmostColumn` | Yes | Zero-based first visible column in the scrollable pane. |
| `topRow` | Yes | Zero-based first visible row in the scrollable pane. |

`pane.type = "SPLIT"` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `xSplitPosition` | Yes | Horizontal split offset in Excel split-pane units (`1/20` of a point). |
| `ySplitPosition` | Yes | Vertical split offset in Excel split-pane units (`1/20` of a point). |
| `leftmostColumn` | Yes | Zero-based first visible column in the right pane. |
| `topRow` | Yes | Zero-based first visible row in the bottom pane. |
| `activePane` | Yes | One of `UPPER_LEFT`, `UPPER_RIGHT`, `LOWER_LEFT`, or `LOWER_RIGHT`. |

Constraints:

- `FROZEN.splitColumn` and `FROZEN.splitRow` must not both be `0`.
- If `FROZEN.splitColumn` is `0`, `FROZEN.leftmostColumn` must also be `0`.
- If `FROZEN.splitRow` is `0`, `FROZEN.topRow` must also be `0`.
- When `FROZEN.splitColumn > 0`, `FROZEN.leftmostColumn` must be greater than or equal to
  `splitColumn`.
- When `FROZEN.splitRow > 0`, `FROZEN.topRow` must be greater than or equal to `splitRow`.
- `SPLIT.xSplitPosition` and `SPLIT.ySplitPosition` must not both be `0`.
- If `SPLIT.xSplitPosition` is `0`, `SPLIT.leftmostColumn` must also be `0`.
- If `SPLIT.ySplitPosition` is `0`, `SPLIT.topRow` must also be `0`.

### SET_SHEET_ZOOM

Set one explicit zoom percentage for a sheet.

```json
{
  "type": "SET_SHEET_ZOOM",
  "sheetName": "Inventory",
  "zoomPercent": 125
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `zoomPercent` | Yes | Integer zoom percentage. Must be between `10` and `400` inclusive. |

### SET_PRINT_LAYOUT

Apply one authoritative supported print-layout state to a sheet. Omitted nested fields normalize
to default or cleared state.

```json
{
  "type": "SET_PRINT_LAYOUT",
  "sheetName": "Inventory",
  "printLayout": {
    "printArea": { "type": "RANGE", "range": "A1:F40" },
    "orientation": "LANDSCAPE",
    "scaling": { "type": "FIT", "widthPages": 1, "heightPages": 0 },
    "repeatingRows": { "type": "BAND", "firstRowIndex": 0, "lastRowIndex": 1 },
    "repeatingColumns": { "type": "BAND", "firstColumnIndex": 0, "lastColumnIndex": 0 },
    "header": { "left": "Inventory", "center": "Q2", "right": "Internal" },
    "footer": { "left": "", "center": "Prepared by GridGrind", "right": "Page 1" },
    "setup": {
      "margins": {
        "left": 0.35,
        "right": 0.55,
        "top": 0.6,
        "bottom": 0.45,
        "header": 0.3,
        "footer": 0.3
      },
      "horizontallyCentered": true,
      "verticallyCentered": true,
      "paperSize": 8,
      "draft": true,
      "blackAndWhite": true,
      "copies": 2,
      "useFirstPageNumber": true,
      "firstPageNumber": 4,
      "rowBreaks": [6],
      "columnBreaks": [3]
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `printLayout` | Yes | Print-layout payload. |

`printLayout` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `printArea` | No | `NONE` or `RANGE`. Defaults to `NONE`. |
| `orientation` | No | `PORTRAIT` or `LANDSCAPE`. Defaults to `PORTRAIT`. |
| `scaling` | No | `AUTOMATIC` or `FIT`. Defaults to `AUTOMATIC`. |
| `repeatingRows` | No | `NONE` or `BAND`. Defaults to `NONE`. |
| `repeatingColumns` | No | `NONE` or `BAND`. Defaults to `NONE`. |
| `header` | No | Plain `left`, `center`, and `right` text segments. Defaults to blank text. |
| `footer` | No | Plain `left`, `center`, and `right` text segments. Defaults to blank text. |
| `setup` | No | Advanced page-setup payload. Defaults to the supported Excel baseline for margins, centering, copies, and page numbering. |

Nested constraints:

- `printArea.type = "RANGE"` requires a non-blank rectangular A1-style range.
- `scaling.type = "FIT"` requires non-negative `widthPages` and `heightPages`; `0` leaves one
  axis unconstrained.
- `repeatingRows.type = "BAND"` requires a zero-based inclusive row band with
  `lastRowIndex >= firstRowIndex` and `lastRowIndex <= 1048575`.
- `repeatingColumns.type = "BAND"` requires a zero-based inclusive column band with
  `lastColumnIndex >= firstColumnIndex` and `lastColumnIndex <= 16383`.

`printLayout.setup` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `margins` | No | `left`, `right`, `top`, `bottom`, `header`, and `footer` margins in inches. Each value must be finite and non-negative. |
| `horizontallyCentered` | No | Center printed content horizontally on the page. Defaults to `false`. |
| `verticallyCentered` | No | Center printed content vertically on the page. Defaults to `false`. |
| `paperSize` | No | Excel paper-size code. Defaults to `0` when omitted. |
| `draft` | No | Draft-print flag. Defaults to `false`. |
| `blackAndWhite` | No | Black-and-white print flag. Defaults to `false`. |
| `copies` | No | Explicit copy count. Defaults to `0` when omitted. |
| `useFirstPageNumber` | No | Whether `firstPageNumber` is authoritative. Defaults to `false`. |
| `firstPageNumber` | No | Explicit first page number. Defaults to `0` when omitted. |
| `rowBreaks` | No | Ordered list of zero-based row indexes for explicit row breaks. Defaults to `[]`. |
| `columnBreaks` | No | Ordered list of zero-based column indexes for explicit column breaks. Defaults to `[]`. |

### CLEAR_PRINT_LAYOUT

Clear the supported print-layout state from a sheet and restore the default supported state.

```json
{
  "type": "CLEAR_PRINT_LAYOUT",
  "sheetName": "Inventory"
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |

---

### SET_CELL

Write a typed value to a single cell using A1 notation. Existing style, hyperlink, and comment
state on the targeted cell is preserved.

```json
{
  "type": "SET_CELL",
  "sheetName": "Inventory",
  "address": "B4",
  "value": { "type": "FORMULA", "formula": "SUM(B2:B3)" }
}
```

FORMULA note: a leading `=` is accepted and stripped automatically. `"=SUM(B2:B3)"` and
`"SUM(B2:B3)"` are equivalent.

DATE / DATE_TIME note: the required Excel number format is applied without discarding any existing
fill, border, font, alignment, or wrap state already present on the cell.

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `address` | Yes | A1-notation cell address. |
| `value` | Yes | Typed cell value. |

---

### SET_RANGE

Write a rectangular grid of typed values. The `rows` matrix must exactly match the dimensions of
`range`. Row count must equal `range` row span; column count of each row must equal `range`
column span. Existing style, hyperlink, and comment state on the targeted cells is preserved.
Any typed value variant accepted by `SET_CELL`, including `RICH_TEXT`, is valid inside `rows`.

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
- `{"type":"FILE","path":"shared/report.xlsx"}`
- `{"type":"DOCUMENT","target":"Inventory!B4"}`

`FILE.path` accepts either a plain relative or absolute file path, or a `file:` URI. Read-side
metadata always returns a normalized plain path string in `path`, never a `file:` URI. Relative
`FILE` targets are analyzed against the workbook's persisted path when one exists, so use
absolute paths when you want cwd-independent health checks.

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

Attach or replace one comment on one cell. The cell is created if needed. GridGrind always
requires the authoritative plain `text` plus `author`, and can also author ordered rich-text
`runs` and an explicit comment `anchor`.

```json
{
  "type": "SET_COMMENT",
  "sheetName": "Inventory",
  "address": "B4",
  "comment": {
    "text": "Lead review scheduled.",
    "author": "GridGrind",
    "visible": false,
    "runs": [
      {
        "text": "Lead",
        "font": {
          "bold": true,
          "fontColorTheme": 4,
          "fontColorTint": -0.2
        }
      },
      { "text": " " },
      {
        "text": "review scheduled.",
        "font": {
          "italic": true,
          "fontColorIndexed": 17
        }
      }
    ],
    "anchor": {
      "firstColumn": 4,
      "firstRow": 1,
      "lastColumn": 8,
      "lastRow": 7
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `address` | Yes | A1-notation cell address. |
| `comment` | Yes | Comment payload. |

`comment.visible` defaults to `false` when omitted.

`comment` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `text` | Yes | Authoritative plain text of the comment. Must be nonblank. |
| `author` | Yes | Nonblank comment author. |
| `visible` | No | Whether the comment box is shown by default. Defaults to `false`. |
| `runs` | No | Ordered rich-text run list. When present, the concatenated run text must equal `comment.text` exactly. |
| `anchor` | No | Explicit comment-box bounds using zero-based `firstColumn`, `firstRow`, `lastColumn`, and `lastRow`. |

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

### SET_PICTURE

Create or replace one named picture-backed drawing object on one sheet. Authored drawing mutations
currently accept only `anchor.type = "TWO_CELL"` with zero-based `from` and `to` markers.

```json
{
  "type": "SET_PICTURE",
  "sheetName": "Ops",
  "picture": {
    "name": "OpsPicture",
    "image": {
      "format": "PNG",
      "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
    },
    "anchor": {
      "type": "TWO_CELL",
      "from": { "columnIndex": 0, "rowIndex": 4, "dx": 0, "dy": 0 },
      "to": { "columnIndex": 2, "rowIndex": 8, "dx": 0, "dy": 0 },
      "behavior": "MOVE_AND_RESIZE"
    },
    "description": "Queue preview"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `picture` | Yes | Authoritative picture payload. |

`picture` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. Replaces any existing drawing object of the same name. |
| `image` | Yes | Binary picture payload with `format` plus base64 bytes. |
| `anchor` | Yes | Authored drawing anchor. Phase 5 supports only `TWO_CELL`. |
| `description` | No | Optional nonblank descriptive text stored on the picture. |

`image.format` accepts `EMF`, `WMF`, `PICT`, `JPEG`, `PNG`, or `DIB`.

`anchor` currently uses:

```json
{
  "type": "TWO_CELL",
  "from": { "columnIndex": 0, "rowIndex": 4, "dx": 0, "dy": 0 },
  "to": { "columnIndex": 2, "rowIndex": 8, "dx": 0, "dy": 0 },
  "behavior": "MOVE_AND_RESIZE"
}
```

`from` and `to` markers use zero-based cell indexes plus non-negative raw cell-relative offsets in
`dx` and `dy`. `behavior` accepts `MOVE_AND_RESIZE`, `MOVE_DONT_RESIZE`, or
`DONT_MOVE_AND_RESIZE` and defaults to `MOVE_AND_RESIZE` when omitted.

---

### SET_SHAPE

Create or replace one named simple shape or connector on one sheet.

```json
{
  "type": "SET_SHAPE",
  "sheetName": "Ops",
  "shape": {
    "name": "OpsShape",
    "kind": "SIMPLE_SHAPE",
    "anchor": {
      "type": "TWO_CELL",
      "from": { "columnIndex": 3, "rowIndex": 4, "dx": 0, "dy": 0 },
      "to": { "columnIndex": 5, "rowIndex": 7, "dx": 0, "dy": 0 },
      "behavior": "MOVE_DONT_RESIZE"
    },
    "presetGeometryToken": "roundRect",
    "text": "Queue"
  }
}
```

```json
{
  "type": "SET_SHAPE",
  "sheetName": "Ops",
  "shape": {
    "name": "OpsConnector",
    "kind": "CONNECTOR",
    "anchor": {
      "type": "TWO_CELL",
      "from": { "columnIndex": 3, "rowIndex": 8, "dx": 0, "dy": 0 },
      "to": { "columnIndex": 6, "rowIndex": 9, "dx": 0, "dy": 0 }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `shape` | Yes | Authoritative shape payload. |

`shape` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. |
| `kind` | Yes | `SIMPLE_SHAPE` or `CONNECTOR`. |
| `anchor` | Yes | Authored drawing anchor. Phase 5 supports only `TWO_CELL`. |
| `presetGeometryToken` | No | Optional preset geometry token for `SIMPLE_SHAPE`. Defaults to `rect` when omitted. Not allowed for `CONNECTOR`. |
| `text` | No | Optional nonblank text for `SIMPLE_SHAPE`. Not allowed for `CONNECTOR`. |

Validation failures are non-mutating. If `presetGeometryToken` is unsupported, GridGrind rejects
the request without deleting an existing object of the same name and without leaving a partial new
shape behind.

---

### SET_EMBEDDED_OBJECT

Create or replace one named embedded object with an authoritative package payload and preview
image.

```json
{
  "type": "SET_EMBEDDED_OBJECT",
  "sheetName": "Ops",
  "embeddedObject": {
    "name": "OpsEmbed",
    "label": "Ops payload",
    "fileName": "ops-payload.txt",
    "command": "open",
    "base64Data": "R3JpZEdyaW5kIGVtYmVkZGVkIHBheWxvYWQ=",
    "previewImage": {
      "format": "PNG",
      "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
    },
    "anchor": {
      "type": "TWO_CELL",
      "from": { "columnIndex": 6, "rowIndex": 4, "dx": 0, "dy": 0 },
      "to": { "columnIndex": 8, "rowIndex": 9, "dx": 0, "dy": 0 }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `embeddedObject` | Yes | Authoritative embedded-object payload. |

`embeddedObject` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. |
| `label` | Yes | Nonblank display label stored on the embedded object. |
| `fileName` | Yes | Nonblank embedded file name surfaced again in payload reads when available. |
| `command` | Yes | Nonblank command string stored on the object metadata. |
| `base64Data` | Yes | Base64-encoded embedded payload bytes. |
| `previewImage` | Yes | Preview picture payload shown by Excel for the object. |
| `anchor` | Yes | Authored drawing anchor. Phase 5 supports only `TWO_CELL`. |

---

### SET_CHART

Create or mutate one named supported simple chart on one sheet. Supported authored families are
`BAR`, `LINE`, and `PIE`. Series bind to workbook ranges or defined names. Existing unsupported
chart detail is preserved on unrelated edits and rejected for authoritative mutation.

```json
{
  "type": "SET_CHART",
  "sheetName": "Ops",
  "chart": {
    "type": "BAR",
    "name": "OpsChart",
    "anchor": {
      "type": "TWO_CELL",
      "from": { "columnIndex": 4, "rowIndex": 0, "dx": 0, "dy": 0 },
      "to": { "columnIndex": 8, "rowIndex": 12, "dx": 0, "dy": 0 },
      "behavior": "MOVE_AND_RESIZE"
    },
    "title": { "type": "TEXT", "text": "Roadmap" },
    "legend": { "type": "VISIBLE", "position": "TOP_RIGHT" },
    "displayBlanksAs": "SPAN",
    "plotOnlyVisibleCells": false,
    "varyColors": true,
    "barDirection": "COLUMN",
    "series": [
      {
        "title": { "type": "TEXT", "text": "Plan" },
        "categories": { "formula": "ChartCategories" },
        "values": { "formula": "Ops!$B$2:$B$4" }
      },
      {
        "title": { "type": "TEXT", "text": "Actual" },
        "categories": { "formula": "ChartCategories" },
        "values": { "formula": "ChartActual" }
      }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `chart` | Yes | Authoritative supported-chart payload. |

`chart` common fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `type` | Yes | `BAR`, `LINE`, or `PIE`. |
| `name` | Yes | Nonblank sheet-local chart name. Reuses the existing chart when the name already exists. |
| `anchor` | Yes | Authored chart-frame anchor. Currently only `TWO_CELL` is supported. |
| `title` | No | `NONE`, `TEXT`, or `FORMULA`. Defaults to `NONE`. |
| `legend` | No | `HIDDEN` or `VISIBLE`. `VISIBLE.position` defaults to `RIGHT` when omitted. |
| `displayBlanksAs` | No | `GAP`, `SPAN`, or `ZERO`. Defaults to `GAP`. |
| `plotOnlyVisibleCells` | No | Whether hidden cells are ignored. Defaults to `true`. |
| `varyColors` | No | Whether Excel varies series colors automatically. Defaults to `false`. |
| `series` | Yes | Ordered non-empty series list. |

`series` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `title` | No | `NONE`, `TEXT`, or `FORMULA`. Defaults to `NONE`. |
| `categories` | Yes | Workbook-bound category source formula. |
| `values` | Yes | Workbook-bound value source formula. |

Variant-specific fields:

| Chart type | Field | Required | Description |
|:-----------|:------|:---------|:------------|
| `BAR` | `barDirection` | No | `COLUMN` or `BAR`. Defaults to `COLUMN`. |
| `PIE` | `firstSliceAngle` | No | Integer between `0` and `360`. |

`categories.formula` and `values.formula` accept either contiguous A1-style references such as
`Ops!$B$2:$B$4` or defined names such as `ChartActual`.
Formula-backed chart titles and series titles must resolve to one cell, either directly or
through a defined name that resolves to one cell.
Validation failures are non-mutating: GridGrind rejects invalid authored chart payloads without
creating a partial chart or partially mutating an existing supported chart.

---

### SET_DRAWING_OBJECT_ANCHOR

Move one existing named drawing object by replacing its anchor authoritatively.

```json
{
  "type": "SET_DRAWING_OBJECT_ANCHOR",
  "sheetName": "Ops",
  "objectName": "OpsPicture",
  "anchor": {
    "type": "TWO_CELL",
    "from": { "columnIndex": 1, "rowIndex": 5, "dx": 0, "dy": 0 },
    "to": { "columnIndex": 3, "rowIndex": 9, "dx": 0, "dy": 0 }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `objectName` | Yes | Existing sheet-local drawing-object name. |
| `anchor` | Yes | Replacement authored drawing anchor. Phase 5 supports only `TWO_CELL`. |

---

### DELETE_DRAWING_OBJECT

Delete one existing named drawing object from one sheet.

```json
{ "type": "DELETE_DRAWING_OBJECT", "sheetName": "Ops", "objectName": "OpsConnector" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `objectName` | Yes | Existing sheet-local drawing-object name to delete. |

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
    "alignment": {
      "wrapText": true,
      "horizontalAlignment": "CENTER",
      "verticalAlignment": "CENTER",
      "textRotation": 15
    },
    "font": {
      "bold": true,
      "fontName": "Aptos",
      "fontHeight": { "type": "POINTS", "points": 13 },
      "fontColor": "#FFFFFF"
    },
    "fill": {
      "pattern": "THIN_HORIZONTAL_BANDS",
      "foregroundColor": "#1F4E78",
      "backgroundColor": "#D9E2F3"
    },
    "border": {
      "all": { "style": "THIN", "color": "#FFFFFF" },
      "bottom": { "style": "DOUBLE", "color": "#D9E2F3" }
    },
    "protection": {
      "locked": true
    }
  }
}
```

```json
{
  "type": "APPLY_STYLE",
  "sheetName": "Inventory",
  "range": "J2:J3",
  "style": {
    "font": {
      "fontColorTheme": 6,
      "fontColorTint": -0.35
    },
    "fill": {
      "gradient": {
        "type": "LINEAR",
        "degree": 45.0,
        "stops": [
          { "position": 0.0, "color": { "rgb": "#1F497D" } },
          { "position": 1.0, "color": { "theme": 4, "tint": 0.45 } }
        ]
      }
    },
    "border": {
      "bottom": {
        "style": "THIN",
        "colorIndexed": 8
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Target sheet. |
| `range` | Yes | Range to style. |
| `style` | Yes | Partial style patch. `numberFormat` plus nested `alignment`, `font`, `fill`, `border`, and `protection` groups are all optional. |

Style fields:

| Field | Type | Values |
|:------|:-----|:-------|
| `numberFormat` | string | Any Excel number format string, e.g. `"#,##0.00"`, `"yyyy-mm-dd"` |
| `alignment` | object | Optional alignment patch. See below. |
| `font` | object | Optional font patch. See below. |
| `fill` | object | Optional fill patch. See below. |
| `border` | object | Optional border patch. See below. |
| `protection` | object | Optional cell-protection patch. See below. |

Alignment patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `alignment.wrapText` | boolean | `true` / `false` |
| `alignment.horizontalAlignment` | string | `"LEFT"`, `"CENTER"`, `"RIGHT"`, `"GENERAL"` |
| `alignment.verticalAlignment` | string | `"TOP"`, `"CENTER"`, `"BOTTOM"` |
| `alignment.textRotation` | integer | Explicit `0..180` text-rotation degrees |
| `alignment.indentation` | integer | Excel's `0..250` cell-indent scale |

Font patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `font.bold` | boolean | `true` / `false` |
| `font.italic` | boolean | `true` / `false` |
| `font.fontName` | string | Excel font family name, e.g. `"Aptos"` |
| `font.fontHeight` | object | Typed font height. See below. |
| `font.fontColor` | string | RGB hex string in `#RRGGBB` form. Lowercase input is normalized to uppercase. |
| `font.fontColorTheme` | integer | Theme color index for the font color. |
| `font.fontColorIndexed` | integer | Indexed color slot for the font color. |
| `font.fontColorTint` | number | Optional tint applied to the chosen font color base. Requires `fontColor`, `fontColorTheme`, or `fontColorIndexed`. |
| `font.underline` | boolean | `true` adds a single underline, `false` removes it |
| `font.strikeout` | boolean | `true` / `false` |

Font-height input:

| Field | Type | Values |
|:------|:-----|:-------|
| `font.fontHeight.type` | string | `"POINTS"` or `"TWIPS"` |
| `font.fontHeight.points` | number | Required when `type` is `"POINTS"`. Must resolve exactly to whole twips, e.g. `11`, `11.5`, `13.25` |
| `font.fontHeight.twips` | integer | Required when `type` is `"TWIPS"`. Positive exact twip value where `20` twips = `1` point |

Structured color notes:

- Color-bearing write fields use one base color family plus an optional tint:
  - RGB: `font.fontColor`, `fill.foregroundColor`, `fill.backgroundColor`, `border.*.color`
  - theme: `font.fontColorTheme`, `fill.foregroundColorTheme`, `fill.backgroundColorTheme`,
    `border.*.colorTheme`
  - indexed: `font.fontColorIndexed`, `fill.foregroundColorIndexed`,
    `fill.backgroundColorIndexed`, `border.*.colorIndexed`
- Lowercase RGB hex input is accepted and normalized to uppercase.
- A tint field requires one corresponding RGB/theme/indexed base field.
- Alpha channels are not part of the public contract.

Fill patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `fill.pattern` | string | `ExcelFillPattern` value such as `"SOLID"`, `"THIN_HORIZONTAL_BANDS"`, `"SQUARES"` |
| `fill.foregroundColor` | string | Foreground RGB color in `#RRGGBB` form |
| `fill.foregroundColorTheme` | integer | Foreground theme color index |
| `fill.foregroundColorIndexed` | integer | Foreground indexed color slot |
| `fill.foregroundColorTint` | number | Optional tint for the chosen foreground base color |
| `fill.backgroundColor` | string | Background RGB color in `#RRGGBB` form for patterned fills |
| `fill.backgroundColorTheme` | integer | Background theme color index for patterned fills |
| `fill.backgroundColorIndexed` | integer | Background indexed color slot for patterned fills |
| `fill.backgroundColorTint` | number | Optional tint for the chosen background base color |
| `fill.gradient` | object | Gradient fill payload. When present, pattern/foreground/background fields must be omitted. |

Fill notes:

- `fill.pattern="NONE"` does not allow colors.
- `fill.pattern="SOLID"` uses `foregroundColor` only; `backgroundColor` is not allowed.
- Patterned fills beyond solid are part of the public contract.
- Gradient fills are authored through `fill.gradient` and are mutually exclusive with patterned
  fill fields.

`fill.gradient` fields:

| Field | Type | Values |
|:------|:-----|:-------|
| `type` | string | Gradient type, e.g. `"LINEAR"` or `"PATH"`. Defaults to `"LINEAR"` when omitted. |
| `degree` | number | Optional angle for linear gradients. |
| `left` | number | Optional left offset for path gradients. |
| `right` | number | Optional right offset for path gradients. |
| `top` | number | Optional top offset for path gradients. |
| `bottom` | number | Optional bottom offset for path gradients. |
| `stops` | array | Ordered gradient stops; each stop has `position` in `0.0..1.0` and a structured `color` object using `rgb`, `theme`, `indexed`, and optional `tint`. |

Border patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `all` | object | Optional default for every border side not explicitly set in the same patch |
| `top` | object | Optional side-specific override |
| `right` | object | Optional side-specific override |
| `bottom` | object | Optional side-specific override |
| `left` | object | Optional side-specific override |

Each border-side object can set a visible style, an RGB color, or both:

| Field | Type | Values |
|:------|:-----|:-------|
| `style` | string | `"NONE"`, `"THIN"`, `"MEDIUM"`, `"DASHED"`, `"DOTTED"`, `"THICK"`, `"DOUBLE"`, `"HAIR"`, `"MEDIUM_DASHED"`, `"DASH_DOT"`, `"MEDIUM_DASH_DOT"`, `"DASH_DOT_DOT"`, `"MEDIUM_DASH_DOT_DOT"`, `"SLANTED_DASH_DOT"` |
| `color` | string | Optional RGB color in `#RRGGBB` form |
| `colorTheme` | integer | Optional theme color index |
| `colorIndexed` | integer | Optional indexed color slot |
| `colorTint` | number | Optional tint applied to the chosen border-color base |

Border notes:

- `border` must set at least one of `all`, `top`, `right`, `bottom`, or `left`.
- `border.all` acts as the default style and color for every side not explicitly set in the same patch.
- Explicit side settings override `border.all`.
- A border color requires an effective visible style on that side, either set on the side itself
  or inherited from `border.all`.
- `style="NONE"` does not allow a `color`.
- `colorTint` requires `color`, `colorTheme`, or `colorIndexed`.

Protection patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `protection.locked` | boolean | `true` / `false` |
| `protection.hiddenFormula` | boolean | `true` / `false` |

Protection note:

- These flags matter when sheet protection is enabled.
- Cell snapshot reads (GET_CELLS, GET_WINDOW, GET_SHEET_SCHEMA) report `style.font.fontHeight`
  as a plain object with both `twips` and `points` fields:
  `{"twips": 260, "points": 13}`. This differs from the write-side
  `FontHeightInput` discriminated format.

---

### SET_DATA_VALIDATION

Create or replace one data-validation rule over a rectangular A1-style range. If the new rule
overlaps existing validation coverage, GridGrind normalizes the sheet so the written rule becomes
authoritative on its target cells and any remaining fragments are retained around it.

```json
{
  "type": "SET_DATA_VALIDATION",
  "sheetName": "Inventory",
  "range": "B2:B200",
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
```

```json
{
  "type": "SET_DATA_VALIDATION",
  "sheetName": "Inventory",
  "range": "C2:C200",
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

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `range` | Yes | Rectangular A1-style target range. |
| `validation` | Yes | Supported data-validation definition. |

Supported rule families:

| `validation.rule.type` | Required fields | Notes |
|:-----------------------|:----------------|:------|
| `EXPLICIT_LIST` | `values` | One or more allowed strings. |
| `FORMULA_LIST` | `formula` | Formula-driven list. A leading `=` is accepted and stripped automatically. |
| `WHOLE_NUMBER` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `DECIMAL_NUMBER` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `DATE` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `TIME` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `TEXT_LENGTH` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `CUSTOM_FORMULA` | `formula` | A leading `=` is accepted and stripped automatically. |

Optional validation fields:

| Field | Description |
|:------|:------------|
| `allowBlank` | Allow empty cells to bypass validation. Defaults to `false`. |
| `suppressDropDownArrow` | Hide Excel's list dropdown arrow when the rule supports it. Defaults to `false`. |
| `prompt` | Optional prompt box with `title`, `text`, and optional `showPromptBox` (defaults to `true`). |
| `errorAlert` | Optional error box with `style`, `title`, `text`, and optional `showErrorBox` (defaults to `true`). |

Comparison operators use the typed string values:
`BETWEEN`, `NOT_BETWEEN`, `EQUAL`, `NOT_EQUAL`, `GREATER_THAN`, `LESS_THAN`,
`GREATER_OR_EQUAL`, and `LESS_OR_EQUAL`.

---

### CLEAR_DATA_VALIDATIONS

Remove data-validation coverage from one sheet. `ALL` clears every validation structure on the
sheet. `SELECTED` removes only the portions whose stored ranges intersect the supplied A1-style
ranges, preserving any remaining coverage fragments around the cleared area.

```json
{
  "type": "CLEAR_DATA_VALIDATIONS",
  "sheetName": "Inventory",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "CLEAR_DATA_VALIDATIONS",
  "sheetName": "Inventory",
  "selection": {
    "type": "SELECTED",
    "ranges": ["B4", "C8:D10"]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `selection` | Yes | `ALL` or an ordered `SELECTED` range list. |

---

### SET_CONDITIONAL_FORMATTING

Creates or replaces one logical conditional-formatting block on the sheet. The block owns one or
more target ranges and one ordered rule list. Authoring supports six rule families:

- `FORMULA_RULE`
- `CELL_VALUE_RULE`
- `COLOR_SCALE_RULE`
- `DATA_BAR_RULE`
- `ICON_SET_RULE`
- `TOP10_RULE`

For `FORMULA_RULE`, `CELL_VALUE_RULE`, and `TOP10_RULE`, the optional `style` payload is
differential, not a whole-cell style patch. Supported differential-style attributes are
`numberFormat`, `bold`, `italic`, `fontHeight`, `fontColor`, `underline`, `strikeout`,
`fillColor`, and per-side differential borders.

```json
{
  "type": "SET_CONDITIONAL_FORMATTING",
  "sheetName": "Inventory",
  "conditionalFormatting": {
    "ranges": ["D2:D200"],
    "rules": [
      {
        "type": "CELL_VALUE_RULE",
        "operator": "GREATER_THAN",
        "formula1": "8",
        "stopIfTrue": false,
        "style": {
          "numberFormat": "0.0",
          "fillColor": "#FDE9D9",
          "fontColor": "#9C0006",
          "bold": true
        }
      }
    ]
  }
}
```

```json
{
  "type": "SET_CONDITIONAL_FORMATTING",
  "sheetName": "Inventory",
  "conditionalFormatting": {
    "ranges": ["K2:K200"],
    "rules": [
      {
        "type": "COLOR_SCALE_RULE",
        "stopIfTrue": false,
        "thresholds": [
          { "type": "MIN" },
          { "type": "PERCENTILE", "value": 50.0 },
          { "type": "MAX" }
        ],
        "colors": [
          { "rgb": "#AA2211" },
          { "rgb": "#FFDD55" },
          { "rgb": "#11CC66" }
        ]
      }
    ]
  }
}
```

Rule family summary:

| Rule type | Required fields | Notes |
|:----------|:----------------|:------|
| `FORMULA_RULE` | `formula` | Optional `style`. |
| `CELL_VALUE_RULE` | `operator`, `formula1` | Optional `formula2` for between/not-between operators and optional `style`. |
| `COLOR_SCALE_RULE` | `thresholds`, `colors` | Threshold and color list sizes must match and must contain at least two control points. |
| `DATA_BAR_RULE` | `color`, `widthMin`, `widthMax`, `minThreshold`, `maxThreshold` | `iconOnly` controls whether only the bar glyph is shown. No direction field is exposed. |
| `ICON_SET_RULE` | `iconSet`, `thresholds` | Threshold count must match the chosen icon-set family. |
| `TOP10_RULE` | `rank` | Optional `style`; `percent` and `bottom` control Top/Bottom N vs percent behavior. |

### CLEAR_CONDITIONAL_FORMATTING

Removes conditional-formatting blocks on the sheet that intersect the provided range selection.
`ALL` clears the sheet. `SELECTED` removes every stored block whose target ranges intersect any of
the selected A1 ranges.

```json
{
  "type": "CLEAR_CONDITIONAL_FORMATTING",
  "sheetName": "Inventory",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "CLEAR_CONDITIONAL_FORMATTING",
  "sheetName": "Inventory",
  "selection": {
    "type": "SELECTED",
    "ranges": ["D2:D200", "F2:F20"]
  }
}
```

---

### SET_AUTOFILTER

Create or replace one sheet-level autofilter range. The range must be rectangular, include a
nonblank header row, and must not overlap any existing table range on the same sheet. Optional
`criteria` author persisted filter-column rules; optional `sortState` authors persisted sort-state
metadata on the same autofilter.

```json
{
  "type": "SET_AUTOFILTER",
  "sheetName": "Inventory",
  "range": "A1:C200",
  "criteria": [
    {
      "columnId": 2,
      "showButton": true,
      "criterion": {
        "type": "VALUES",
        "values": ["Queued", "Done"],
        "includeBlank": false
      }
    }
  ],
  "sortState": {
    "range": "A2:C200",
    "caseSensitive": false,
    "columnSort": false,
    "sortMethod": "",
    "conditions": [
      {
        "range": "C2:C200",
        "descending": true,
        "sortBy": ""
      }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |
| `range` | Yes | Rectangular A1-style range. The first row is treated as the filter header row. |
| `criteria` | No | Ordered authored filter-column list. Defaults to `[]`. |
| `sortState` | No | Authored sort-state payload. Omit it to leave the autofilter without persisted sort metadata. |

If a sheet already has a sheet-level autofilter, the new range replaces it.

`criteria[*]` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `columnId` | Yes | Zero-based column offset within the autofilter range. |
| `showButton` | No | Whether Excel shows the dropdown button for that filter column. Defaults to `true`. |
| `criterion` | Yes | One authored criterion payload. |

`criterion` variants:

- `VALUES`: `values` plus `includeBlank`
- `CUSTOM`: `and` plus one or more `{ "operator": "...", "value": "..." }` conditions
- `DYNAMIC`: Excel dynamic-filter token in `type` plus optional numeric `value` and `maxValue`
- `TOP10`: `value`, `top`, and `percent`
- `COLOR`: `cellColor` plus a structured `color` object using `rgb`, `theme`, `indexed`, and optional `tint`
- `ICON`: `iconSet` plus `iconId`

`sortState` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `range` | Yes | A1-style range covered by the stored sort state. |
| `caseSensitive` | No | Case-sensitive sort flag. Defaults to `false`. |
| `columnSort` | No | Column-oriented sort flag. Defaults to `false`. |
| `sortMethod` | No | Excel sort-method token. Defaults to `""`. |
| `conditions` | Yes | Ordered list of sort conditions. Must not be empty. |

Each `sortState.conditions[*]` entry carries `range`, `descending`, optional `sortBy`, optional
structured `color`, and optional `iconId`. Use blank `sortBy` for ordinary value sorts.

---

### CLEAR_AUTOFILTER

Remove the sheet-level autofilter range from one sheet. Table-owned autofilters remain attached to
their tables.

```json
{ "type": "CLEAR_AUTOFILTER", "sheetName": "Inventory" }
```

| Field | Required | Description |
|:------|:---------|:------------|
| `sheetName` | Yes | Existing target sheet. |

---

### SET_TABLE

Create or replace one workbook-global table definition. The first row of `table.range` supplies
the table's header cells, which must be nonblank and unique case-insensitively. Same-name tables
on the same sheet are replaced. Same-name tables on a different sheet are rejected. Overlapping
different-name tables are rejected. If the new table overlaps a sheet-level autofilter, the
sheet-level filter is cleared so the table-owned autofilter becomes authoritative on that range.
Later value writes and style patches that touch the table header row keep the persisted
table-column metadata converged with the visible header cells.

```json
{
  "type": "SET_TABLE",
  "table": {
    "name": "InventoryTable",
    "sheetName": "Inventory",
    "range": "A1:C200",
    "style": { "type": "NONE" }
  }
}
```

```json
{
  "type": "SET_TABLE",
  "table": {
    "name": "InventoryTable",
    "sheetName": "Inventory",
    "range": "A1:C200",
    "showTotalsRow": true,
    "hasAutofilter": false,
    "comment": "Inventory tracker",
    "published": true,
    "insertRow": true,
    "insertRowShift": true,
    "headerRowCellStyle": "InventoryHeader",
    "dataCellStyle": "InventoryData",
    "totalsRowCellStyle": "InventoryTotals",
    "style": {
      "type": "NAMED",
      "name": "TableStyleMedium2",
      "showFirstColumn": false,
      "showLastColumn": false,
      "showRowStripes": true,
      "showColumnStripes": false
    },
    "columns": [
      { "columnIndex": 0, "totalsRowLabel": "Total" },
      { "columnIndex": 1, "totalsRowFunction": "SUM" },
      { "columnIndex": 2, "uniqueName": "status-unique" }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `table` | Yes | One workbook-global table definition payload. |

`table` payload:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Workbook-global table identifier. Must be a valid defined-name token and must not conflict with an existing defined name. |
| `sheetName` | Yes | Existing sheet that owns the table. |
| `range` | Yes | Rectangular A1-style table range. Must include at least a header row plus one data row; with `showTotalsRow=true`, it must include one additional totals row. |
| `showTotalsRow` | No | Whether the table includes a totals row. Defaults to `false`. |
| `hasAutofilter` | No | Whether the table owns an autofilter. Defaults to `true`. |
| `style` | Yes | Table style definition: `NONE` or `NAMED`. |
| `comment` | No | Table comment metadata. Defaults to `""`. |
| `published` | No | Published flag. Defaults to `false`. |
| `insertRow` | No | Insert-row flag. Defaults to `false`. |
| `insertRowShift` | No | Insert-row-shift flag. Defaults to `false`. |
| `headerRowCellStyle` | No | Header-row cell-style metadata. Defaults to `""`. |
| `dataCellStyle` | No | Data-cell style metadata. Defaults to `""`. |
| `totalsRowCellStyle` | No | Totals-row cell-style metadata. Defaults to `""`. |
| `columns` | No | Advanced authored column metadata. Defaults to `[]`. |

Omit `showTotalsRow` when the table has no totals row. Setting `"showTotalsRow": false` is
accepted but redundant.

Supported table-style variants:

- `{"type":"NONE"}`
- `{"type":"NAMED","name":"TableStyleMedium2","showFirstColumn":false,"showLastColumn":false,"showRowStripes":true,"showColumnStripes":false}`

Named styles must exist in the workbook style source or the write fails.

`columns[*]` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `columnIndex` | Yes | Zero-based table-column ordinal inside the table range. |
| `uniqueName` | No | Optional unique-name metadata. Defaults to `""`. |
| `totalsRowLabel` | No | Optional totals-row label. Defaults to `""`. |
| `totalsRowFunction` | No | Optional totals-row function token. Input is case-insensitive and canonicalized to Excel's lowercase token family. Defaults to `""`. |
| `calculatedColumnFormula` | No | Optional calculated-column formula. Defaults to `""`. |

---

### DELETE_TABLE

Delete one existing workbook-global table by name and expected sheet.

```json
{
  "type": "DELETE_TABLE",
  "name": "InventoryTable",
  "sheetName": "Inventory"
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Workbook-global table name to delete. |
| `sheetName` | Yes | Sheet that must own the table. |

The request fails if the table does not exist on the expected sheet.

---

### SET_PIVOT_TABLE

Create or replace one workbook-global pivot-table definition to POI XSSF's supported limited
extent. Sources may be an explicit contiguous sheet range, an existing named range, or an existing
table. `rowLabels`, `columnLabels`, `reportFilters`, and `dataFields` must use disjoint source
columns because POI persists only one role per pivot field. When `reportFilters` is non-empty,
`anchor.topLeftAddress` must be on Excel row `3` or lower so Excel's page-filter layout has room
above the rendered body.

```json
{
  "type": "SET_PIVOT_TABLE",
  "pivotTable": {
    "name": "SalesPivot",
    "sheetName": "Report",
    "source": {
      "type": "RANGE",
      "sheetName": "Inventory",
      "range": "A1:D200"
    },
    "anchor": { "topLeftAddress": "C5" },
    "rowLabels": ["Region"],
    "columnLabels": ["Stage"],
    "reportFilters": [],
    "dataFields": [
      {
        "sourceColumnName": "Amount",
        "function": "SUM",
        "displayName": "Total Amount",
        "valueFormat": "#,##0.00"
      }
    ]
  }
}
```

```json
{
  "type": "SET_PIVOT_TABLE",
  "pivotTable": {
    "name": "NamedPivot",
    "sheetName": "NamedReport",
    "source": { "type": "NAMED_RANGE", "name": "PivotSource" },
    "anchor": { "topLeftAddress": "A3" },
    "rowLabels": ["Region"],
    "columnLabels": [],
    "reportFilters": ["Owner"],
    "dataFields": [
      {
        "sourceColumnName": "Amount",
        "function": "SUM"
      }
    ]
  }
}
```

```json
{
  "type": "SET_PIVOT_TABLE",
  "pivotTable": {
    "name": "TablePivot",
    "sheetName": "TableReport",
    "source": { "type": "TABLE", "name": "InventoryTable" },
    "anchor": { "topLeftAddress": "F4" },
    "rowLabels": ["Stage"],
    "columnLabels": [],
    "reportFilters": [],
    "dataFields": [
      {
        "sourceColumnName": "Amount",
        "function": "SUM",
        "displayName": "Total Amount"
      }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `pivotTable` | Yes | One workbook-global pivot-table definition payload. |

`pivotTable` payload:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Workbook-global pivot-table identifier. Must be nonblank and valid for persisted pivot metadata. |
| `sheetName` | Yes | Existing destination sheet that owns the pivot-table relation. |
| `source` | Yes | One authored pivot-source payload: `RANGE`, `NAMED_RANGE`, or `TABLE`. |
| `anchor` | Yes | Top-left destination address for the rendered pivot. When `reportFilters` is non-empty, this must be on Excel row `3` or lower. |
| `rowLabels` | No | Ordered distinct source-column names for row fields. Defaults to `[]`. |
| `columnLabels` | No | Ordered distinct source-column names for column fields. Defaults to `[]`. |
| `reportFilters` | No | Ordered distinct source-column names for page filters. Defaults to `[]`. |
| `dataFields` | Yes | Ordered non-empty data-field list. |

Supported `source` variants:

- `RANGE`: `sheetName` plus contiguous A1 `range`. The first row is the header row.
- `NAMED_RANGE`: existing named-range `name`. The named range may be workbook-scoped or
  sheet-scoped, but it must resolve to one contiguous area.
- `TABLE`: existing workbook-global table `name`.

`dataFields[*]` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `sourceColumnName` | Yes | Source-column header text for the aggregated value field. |
| `function` | Yes | Pivot aggregation function: `SUM`, `COUNT`, `COUNT_NUMS`, `AVERAGE`, `MAX`, `MIN`, `PRODUCT`, `STD_DEV`, `STD_DEVP`, `VAR`, or `VARP`. |
| `displayName` | No | Visible data-field caption. Defaults to `sourceColumnName`. |
| `valueFormat` | No | Optional persisted number-format string. |

---

### DELETE_PIVOT_TABLE

Delete one existing workbook-global pivot table by name and expected sheet.

```json
{
  "type": "DELETE_PIVOT_TABLE",
  "name": "SalesPivot",
  "sheetName": "Report"
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Existing workbook-global pivot-table name. |
| `sheetName` | Yes | Existing sheet expected to own the pivot table. |

The request fails if the pivot table does not exist on the expected sheet.

---

### APPEND_ROW

Append a single row of typed values after the last value-bearing row in a sheet.
If the next append position lands on cells that already exist because they carry only style,
hyperlink, or comment state, that existing state is preserved when values are written there.
Any typed value variant accepted by `SET_CELL`, including `RICH_TEXT`, is valid inside `values`.

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

Rows that contain only style, comment, or hyperlink metadata are ignored when locating the append
position.

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

GridGrind uses deterministic content-based sizing rather than host font metrics, so Docker,
headless, and local runs produce the same widths.

---

### EVALUATE_FORMULAS

Force recalculation of all formulas in the workbook before save. Use this when formulas depend on
data written earlier in the same request.

```json
{ "type": "EVALUATE_FORMULAS" }
```

No additional fields.

---

### EVALUATE_FORMULA_CELLS

Recalculate one or more explicit formula cells and persist only those refreshed cached results.
Use this when a request needs targeted recalculation after a narrow workbook edit.

```json
{
  "type": "EVALUATE_FORMULA_CELLS",
  "cells": [
    { "sheetName": "Budget", "address": "D2" },
    { "sheetName": "Budget", "address": "E2" }
  ]
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `cells` | Yes | Ordered non-empty list of explicit formula-cell targets. |

Each target must point at an existing formula cell.

---

### CLEAR_FORMULA_CACHES

Clear all persisted cached formula results from the workbook and reset the in-process evaluator
state. Later formula reads in the same request still evaluate live, but saved workbooks no longer
carry stale cached values.

```json
{ "type": "CLEAR_FORMULA_CACHES" }
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
sheet-qualified cells or rectangular ranges, or formula-defined name targets.

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

```json
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetRollup",
  "scope": { "type": "WORKBOOK" },
  "target": { "formula": "SUM(Budget!$B$2:$B$5)" }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Defined-name identifier. Must not collide with A1 or R1C1 reference syntax and must not use the reserved `_xlnm.` prefix. |
| `scope` | Yes | Workbook or sheet scope payload. |
| `target` | Yes | Either an explicit sheet-qualified cell/range target or a formula-defined target. Reversed explicit range endpoints are normalized to top-left:`bottom-right`. Formula-defined targets must set `formula` only, not `sheetName`/`range`. |

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

Response shapes:

- `{"kind":"EMPTY","sheetCount":0,"sheetNames":[],"namedRangeCount":0,"forceFormulaRecalculationOnOpen":false}`
- `{"kind":"WITH_SHEETS","sheetCount":2,"sheetNames":["Budget","Budget Review"],"activeSheetName":"Budget Review","selectedSheetNames":["Budget","Budget Review"],"namedRangeCount":0,"forceFormulaRecalculationOnOpen":false}`

`selectedSheetNames` are returned in workbook order, not request order.

### GET_WORKBOOK_PROTECTION

Returns workbook-level protection facts such as structure, windows, and revisions lock state plus
whether the workbook stores password hashes for the workbook or revisions protection domains.

```json
{ "type": "GET_WORKBOOK_PROTECTION", "requestId": "workbook-protection" }
```

Response shape:

```json
{
  "protection": {
    "structureLocked": false,
    "windowsLocked": false,
    "revisionLocked": false,
    "workbookPasswordHashPresent": false,
    "revisionsPasswordHashPresent": false
  }
}
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
| `visibility` | Sheet visibility state: `VISIBLE`, `HIDDEN`, or `VERY_HIDDEN`. |
| `protection.kind` | `UNPROTECTED` or `PROTECTED`. |
| `protection.settings` | Present only when `protection.kind=PROTECTED`; echoes the supported authored lock flags. |
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
| `declaredType` | Raw Excel cell type: `BLANK`, `STRING`, `NUMBER`, `BOOLEAN`, `FORMULA`, or `ERROR`. |
| `effectiveType` | For non-formula cells: same as `declaredType`. For formula cells: always `FORMULA` — the evaluated result type is in `evaluation.effectiveType`. |
| `displayValue` | Formatted string as Excel would render it. |
| `style` | Nested style snapshot with `numberFormat`, `alignment`, `font`, `fill`, `border`, and `protection`. |
| `hyperlink` | Optional hyperlink metadata attached at snapshot time. |
| `comment` | Optional comment metadata attached at snapshot time, including plain text, visibility, optional rich-text runs, and optional anchor bounds when present. |

Type-specific value fields (present only on matching `effectiveType`):

| Field | Present when | Description |
|:------|:-------------|:------------|
| `stringValue` | `effectiveType=STRING` | The string content. |
| `richText` | `effectiveType=STRING` | Optional ordered rich-text runs. When present, run text concatenates exactly to `stringValue`, and each run carries factual effective font data. |
| `numberValue` | `effectiveType=NUMBER` | The numeric value as a double. |
| `booleanValue` | `effectiveType=BOOLEAN` | `true` or `false`. |
| `errorValue` | `effectiveType=ERROR` | The error string (e.g. `#DIV/0!`). |
| `formula` | `declaredType=FORMULA` | The formula text without the leading `=`. |
| `evaluation` | `declaredType=FORMULA` | Nested cell snapshot of the evaluated result. |

**fontHeight read-back shape:** `style.font.fontHeight` is a plain object with both `twips` and
`points` fields: `{"twips": 260, "points": 13}`. This differs from the write-side
`FontHeightInput` format which uses a discriminated `{"type": "POINTS", "points": 13}` object.
Agents that read a font height and want to write it back must use the `FontHeightInput` write
format, not the read-back shape.

**read-side color and gradient shape:** factual read results use structured color objects rather
than raw `#RRGGBB` strings. `style.font.fontColor`, `style.fill.foregroundColor`,
`style.fill.backgroundColor`, and `style.border.*.color` are `CellColorReport` objects with
`rgb` plus optional `theme`, `indexed`, and `tint` facts when Excel stores them. Gradient fills
come back under `style.fill.gradient` with `type`, optional geometry (`degree`, `left`, `right`,
`top`, `bottom`), and ordered `stops` carrying `position` plus structured colors. The write-side
style contract uses asymmetric field names rather than the read-back object shape:
`fontColor`/`fontColorTheme`/`fontColorIndexed`/`fontColorTint`, corresponding
foreground/background/border color fields, and `fill.gradient`. Agents that read a
`CellColorReport` and want to write it back must translate the structured report into those
write-side fields rather than echo the read-back object verbatim.

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

Returns hyperlink metadata for selected cells on one sheet. Response hyperlinks reuse the same
discriminated shape as `SET_HYPERLINK` targets. `FILE` targets come back in the `path` field as
normalized plain path strings, not `file:` URIs.

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

Each returned entry includes the cell `address` plus a `comment` object. `comment.runs` is an
optional ordered rich-text run list whose text concatenates exactly to `comment.text`.
`comment.anchor` is optional and, when present, exposes zero-based comment-box bounds as
`firstColumn`, `firstRow`, `lastColumn`, and `lastRow`.

### GET_DRAWING_OBJECTS

Returns factual drawing-object metadata for one sheet.

```json
{
  "type": "GET_DRAWING_OBJECTS",
  "requestId": "drawing-objects",
  "sheetName": "Ops"
}
```

Response shape: `{ "drawingObjects": [ ... ] }`.

Returned entries are one of:

- `PICTURE` with `format`, `contentType`, byte size or digest facts, optional pixel size, optional
  `description`, and a factual `anchor`
- `CHART` with `supported`, ordered `plotTypeTokens`, title text, and a factual `anchor`
- `SHAPE` with `kind`, optional `presetGeometryToken`, optional `text`, `childCount`, and a
  factual `anchor`
- `EMBEDDED_OBJECT` with `packagingKind`, content type, digest facts, optional label or file name
  or command metadata, optional preview-image facts, and a factual `anchor`

Read-side anchors can be `TWO_CELL`, `ONE_CELL`, or `ABSOLUTE`. `TWO_CELL` markers expose
zero-based `columnIndex`, `rowIndex`, `dx`, and `dy`. `ONE_CELL` and `ABSOLUTE` anchors expose
their size fields in EMUs.

### GET_CHARTS

Returns factual chart metadata for one sheet. Supported simple `BAR`, `LINE`, and `PIE` charts are
modeled authoritatively. Unsupported plot families or multi-plot combinations are surfaced as
explicit `UNSUPPORTED` entries with preserved plot-type tokens.

```json
{
  "type": "GET_CHARTS",
  "requestId": "charts",
  "sheetName": "Ops"
}
```

Response shape: `{ "charts": [ ... ] }`.

Returned entries are one of:

- `BAR` with chart `anchor`, `title`, `legend`, `displayBlanksAs`, `plotOnlyVisibleCells`,
  `varyColors`, `barDirection`, ordered `axes`, and ordered `series`
- `LINE` with chart `anchor`, `title`, `legend`, `displayBlanksAs`, `plotOnlyVisibleCells`,
  `varyColors`, ordered `axes`, and ordered `series`
- `PIE` with chart `anchor`, `title`, `legend`, `displayBlanksAs`, `plotOnlyVisibleCells`,
  `varyColors`, optional `firstSliceAngle`, and ordered `series`
- `UNSUPPORTED` with chart `anchor`, ordered `plotTypeTokens`, and human-readable `detail`

Series titles are returned as `NONE`, `TEXT`, or `FORMULA`. Series data sources are returned as
`STRING_REFERENCE`, `NUMERIC_REFERENCE`, `STRING_LITERAL`, or `NUMERIC_LITERAL`. Read-side anchors
reuse the same factual `TWO_CELL`, `ONE_CELL`, or `ABSOLUTE` shapes described by
`GET_DRAWING_OBJECTS`.
Blank loaded chart titles normalize to `NONE`. Sparse literal source caches preserve declared point
order by surfacing missing positions as empty strings. If a workbook still contains a graphic frame
but its chart relationship is missing, `GET_CHARTS` omits the broken authoritative chart entry and
`GET_DRAWING_OBJECTS` reports the surviving frame as a read-only `GRAPHIC_FRAME` shape instead of
failing the sheet read.

### GET_DRAWING_OBJECT_PAYLOAD

Returns the extracted binary payload for one existing named picture or embedded object on one
sheet.

```json
{
  "type": "GET_DRAWING_OBJECT_PAYLOAD",
  "requestId": "picture-payload",
  "sheetName": "Ops",
  "objectName": "OpsPicture"
}
```

```json
{
  "type": "GET_DRAWING_OBJECT_PAYLOAD",
  "requestId": "embedded-payload",
  "sheetName": "Ops",
  "objectName": "OpsEmbed"
}
```

Response shape: `{ "payload": { ... } }`.

Returned payload entries are:

- `PICTURE` with `format`, `contentType`, `fileName`, `sha256`, `base64Data`, and optional
  `description`
- `EMBEDDED_OBJECT` with `packagingKind`, `contentType`, optional `fileName`, `sha256`,
  `base64Data`, and optional `label` or `command`

Named non-binary shapes such as connectors and simple shapes are rejected because they do not own
an extractable binary payload.

### GET_SHEET_LAYOUT

Returns pane state, effective zoom, and per-row or per-column layout facts for one sheet. Row and
column entries include explicit size plus `hidden`, `outlineLevel`, and `collapsed` where Excel
exposes that state.

```json
{
  "type": "GET_SHEET_LAYOUT",
  "requestId": "layout",
  "sheetName": "Inventory"
}
```

The returned `layout.pane` is one of:

- `NONE`
- `FROZEN` with `splitColumn`, `splitRow`, `leftmostColumn`, and `topRow`
- `SPLIT` with `xSplitPosition`, `ySplitPosition`, `leftmostColumn`, `topRow`, and `activePane`

### GET_PRINT_LAYOUT

Returns the supported print-layout state for one sheet.

```json
{
  "type": "GET_PRINT_LAYOUT",
  "requestId": "print-layout",
  "sheetName": "Inventory"
}
```

The returned `printLayout.setup` object carries advanced page-setup facts: `margins`,
`horizontallyCentered`, `verticallyCentered`, `paperSize`, `draft`, `blackAndWhite`, `copies`,
`useFirstPageNumber`, `firstPageNumber`, and explicit `rowBreaks` plus `columnBreaks`.

### GET_DATA_VALIDATIONS

Returns factual data-validation structures for one sheet. Each returned entry is one of:

- `SUPPORTED`: a fully modeled validation definition plus its normalized stored ranges
- `UNSUPPORTED`: a present workbook rule GridGrind can detect but cannot expose as a supported
  validation definition; the entry includes `kind` and `detail` so callers can distinguish
  unsupported families from invalid workbook structures that Apache POI still exposes

```json
{
  "type": "GET_DATA_VALIDATIONS",
  "requestId": "data-validations",
  "sheetName": "Inventory",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "GET_DATA_VALIDATIONS",
  "requestId": "selected-data-validations",
  "sheetName": "Inventory",
  "selection": {
    "type": "SELECTED",
    "ranges": ["B2:B200", "C2:C200"]
  }
}
```

Range-selection payloads use:

```json
{ "type": "ALL" }
{
  "type": "SELECTED",
  "ranges": ["B2:B200", "C2:C200"]
}
```

### GET_CONDITIONAL_FORMATTING

Returns factual conditional-formatting blocks for one sheet. Each returned block includes its
stored ranges plus an ordered rule list. Rule reports may be one of:

- `FORMULA_RULE`
- `CELL_VALUE_RULE`
- `COLOR_SCALE_RULE`
- `DATA_BAR_RULE`
- `ICON_SET_RULE`
- `UNSUPPORTED_RULE`

Unreadable raw OOXML rule-family metadata degrades to `UNSUPPORTED_RULE` so malformed loaded
workbooks can still be inspected instead of aborting the read.

```json
{
  "type": "GET_CONDITIONAL_FORMATTING",
  "requestId": "conditional-formatting",
  "sheetName": "Inventory",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "GET_CONDITIONAL_FORMATTING",
  "requestId": "selected-conditional-formatting",
  "sheetName": "Inventory",
  "selection": {
    "type": "SELECTED",
    "ranges": ["A2:D200", "F2:F20"]
  }
}
```

### GET_AUTOFILTERS

Returns factual autofilter metadata for one sheet. The result may include:

- `SHEET`: one sheet-owned autofilter stored directly on the worksheet
- `TABLE`: one table-owned autofilter stored on a table definition, including `tableName`

```json
{
  "type": "GET_AUTOFILTERS",
  "requestId": "autofilters",
  "sheetName": "Inventory"
}
```

Each returned autofilter includes its stored `range`, persisted `filterColumns`, and optional
`sortState`. `filterColumns[*].criterion` is one of:

- `VALUES`
- `CUSTOM`
- `DYNAMIC`
- `TOP10`
- `COLOR`
- `ICON`

When present, `sortState` carries `range`, `caseSensitive`, `columnSort`, `sortMethod`, and the
ordered `conditions` Excel stores for the autofilter.

GridGrind reports persisted sort-state and sort-condition ranges exactly as stored, including
blank raw ranges, so malformed workbook metadata is surfaced factually instead of being dropped.

### GET_TABLES

Returns factual table metadata selected by workbook-global table name or all tables.

```json
{
  "type": "GET_TABLES",
  "requestId": "tables",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "GET_TABLES",
  "requestId": "selected-tables",
  "selection": {
    "type": "BY_NAMES",
    "names": ["InventoryTable", "Trips"]
  }
}
```

Each returned table includes:

- `name`
- `sheetName`
- `range`
- `headerRowCount`
- `totalsRowCount`
- `columnNames`
- `columns`
- `style`
- `hasAutofilter`
- `comment`
- `published`
- `insertRow`
- `insertRowShift`
- `headerRowCellStyle`
- `dataCellStyle`
- `totalsRowCellStyle`

Each `columns[*]` entry includes the persisted table-column `id`, visible `name`, `uniqueName`,
`totalsRowLabel`, `totalsRowFunction`, and `calculatedColumnFormula`.

Table-selection payloads use:

```json
{ "type": "ALL" }
{
  "type": "BY_NAMES",
  "names": ["InventoryTable", "Trips"]
}
```

### GET_PIVOT_TABLES

Returns factual pivot-table metadata selected by workbook-global pivot-table name or all pivots.
Supported pivots surface source, stored anchor, row or column labels, report filters, data fields,
and values-axis placement. Unsupported or malformed loaded pivots are returned explicitly with
preserved detail instead of causing read failure.

```json
{
  "type": "GET_PIVOT_TABLES",
  "requestId": "pivots",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "GET_PIVOT_TABLES",
  "requestId": "selected-pivots",
  "selection": {
    "type": "BY_NAMES",
    "names": ["SalesPivot", "NamedPivot"]
  }
}
```

Response shape: `{ "pivotTables": [ ... ] }`.

Returned entries are one of:

- `SUPPORTED` with persisted `source`, stored `anchor`, ordered `rowLabels`, `columnLabels`,
  `reportFilters`, ordered `dataFields`, and `valuesAxisOnColumns`
- `UNSUPPORTED` with stored `anchor` plus human-readable `detail`

Returned `source` variants are:

- `RANGE` with `sheetName` and `range`
- `NAMED_RANGE` with original `name` plus the currently resolved `sheetName` and `range`
- `TABLE` with original `name` plus the currently resolved `sheetName` and `range`

Returned `anchor` includes authored `topLeftAddress` plus the current persisted `locationRange`.
Each returned field entry carries `sourceColumnIndex` and `sourceColumnName`. Data-field entries
also carry `function`, `displayName`, and optional `valueFormat`.

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
for the highest count. Formula cells contribute their evaluated result type (e.g. `NUMBER`,
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

### ANALYZE_DATA_VALIDATION_HEALTH

Reports finding-bearing data-validation health across one or more sheets. Findings include
unsupported rules, broken formulas, and overlapping validation coverage.

```json
{
  "type": "ANALYZE_DATA_VALIDATION_HEALTH",
  "requestId": "data-validation-health",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Inventory", "Summary"]
  }
}
```

### ANALYZE_CONDITIONAL_FORMATTING_HEALTH

Reports conditional-formatting findings such as broken formulas, unsupported loaded rule
families, empty target ranges, or priority collisions.

```json
{
  "type": "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
  "requestId": "conditional-formatting-health",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Inventory", "Summary"]
  }
}
```

### ANALYZE_AUTOFILTER_HEALTH

Reports autofilter findings such as invalid ranges, blank header rows, or ownership mismatches
between sheet-level filters and tables.

```json
{
  "type": "ANALYZE_AUTOFILTER_HEALTH",
  "requestId": "autofilter-health",
  "selection": {
    "type": "SELECTED",
    "sheetNames": ["Inventory", "Summary"]
  }
}
```

### ANALYZE_TABLE_HEALTH

Reports table findings such as overlaps, broken ranges, blank or duplicate headers, or style
mismatches.

```json
{
  "type": "ANALYZE_TABLE_HEALTH",
  "requestId": "table-health",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "ANALYZE_TABLE_HEALTH",
  "requestId": "selected-table-health",
  "selection": {
    "type": "BY_NAMES",
    "names": ["InventoryTable", "Trips"]
  }
}
```

### ANALYZE_PIVOT_TABLE_HEALTH

Reports pivot-table findings such as missing cache parts, missing workbook-cache definitions,
broken sources, duplicate names, synthetic fallback names, or unsupported persisted detail.

```json
{
  "type": "ANALYZE_PIVOT_TABLE_HEALTH",
  "requestId": "pivot-health",
  "selection": { "type": "ALL" }
}
```

```json
{
  "type": "ANALYZE_PIVOT_TABLE_HEALTH",
  "requestId": "selected-pivot-health",
  "selection": {
    "type": "BY_NAMES",
    "names": ["SalesPivot", "NamedPivot"]
  }
}
```

### ANALYZE_HYPERLINK_HEALTH

Reports hyperlink findings such as malformed external targets, missing local file targets,
unresolved relative file targets, or broken document destinations.

Relative `FILE` targets are resolved against the workbook's persisted path when `source` or
`persistence` gives the workbook a filesystem location. When the workbook is still unsaved,
relative `FILE` targets are reported as `HYPERLINK_UNRESOLVED_FILE_TARGET`.

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

Runs every shipped finding-bearing analysis family across the workbook and returns one aggregated
flat finding list. This is the primary workbook-health check and works especially well with
`persistence.type=NONE` when you want a no-save lint pass.

The aggregate currently includes formula, data validation, conditional formatting, autofilter,
table, pivot-table, hyperlink, and named-range findings.

```json
{
  "type": "ANALYZE_WORKBOOK_FINDINGS",
  "requestId": "workbook-findings"
}
```

Formula references to same-request sheet names with spaces should use single quotes, for example
`'Budget Review'!B1`. When execution succeeds, GridGrind reports this style of request issue in
the success response `warnings` array.
