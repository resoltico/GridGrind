---
afad: "3.5"
version: "0.44.0"
domain: QUICK_REFERENCE
updated: "2026-04-13"
route:
  keywords: [gridgrind, quick-reference, snippets, json, operations, reads, introspection, analysis, copy-paste, ensure-sheet, rename-sheet, delete-sheet, move-sheet, copy-sheet, set-active-sheet, set-selected-sheets, set-sheet-visibility, set-sheet-protection, clear-sheet-protection, set-workbook-protection, clear-workbook-protection, merge-cells, unmerge-cells, set-column-width, set-row-height, set-sheet-pane, set-sheet-zoom, set-print-layout, clear-print-layout, freeze-panes, split-panes, set-cell, set-range, set-hyperlink, clear-hyperlink, set-comment, clear-comment, set-picture, set-chart, set-pivot-table, set-shape, set-embedded-object, set-drawing-object-anchor, delete-drawing-object, set-data-validation, clear-data-validations, set-autofilter, clear-autofilter, set-table, delete-table, delete-pivot-table, set-named-range, delete-named-range, apply-style, append-row, clear-range, evaluate-formulas, get-cells, get-window, get-print-layout, get-package-security, get-workbook-protection, get-data-validations, get-autofilters, get-tables, get-pivot-tables, get-drawing-objects, get-charts, get-drawing-object-payload, get-sheet-schema, analyze-autofilter-health, analyze-table-health, analyze-pivot-table-health, analyze-workbook-findings, ooxml, package-security, encryption, signing, coordinates, rowindex, columnindex, warnings]
  questions: ["gridgrind json snippets", "how do I write a cell in gridgrind", "gridgrind copy paste examples", "gridgrind copy sheet example", "gridgrind active sheet example", "gridgrind selected sheets example", "gridgrind sheet visibility example", "gridgrind sheet protection example", "gridgrind workbook protection example", "gridgrind package security example", "how do I open an encrypted workbook in gridgrind", "how do I inspect package signatures in gridgrind", "gridgrind hyperlink example", "gridgrind comment example", "gridgrind picture example", "gridgrind chart example", "gridgrind pivot table example", "how do I read pivot tables in gridgrind", "how do I lint pivot tables in gridgrind", "gridgrind drawing payload example", "gridgrind table example", "gridgrind autofilter example", "gridgrind named range example", "what do gridgrind reads look like", "which gridgrind fields use A1 versus zero-based indexes", "how do I lint workbook health without saving"]
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
  "executionMode": { ... },
  "formulaEnvironment": { ... },
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

`formulaEnvironment` is optional. Use it when evaluation needs external workbook bindings,
cached-value fallback for unresolved external references, or template-backed UDFs.

`source.security.password` is optional and applies only to encrypted `EXISTING` workbook sources.
Use it when the `.xlsx` package is password-protected.

`persistence.security` is optional on `SAVE_AS` and `OVERWRITE`. Use it to encrypt or sign the
persisted `.xlsx` package:

```json
{
  "persistence": {
    "type": "SAVE_AS",
    "path": "secured.xlsx",
    "security": {
      "encryption": { "password": "GridGrind-2026" },
      "signature": {
        "pkcs12Path": "signing-material.p12",
        "keystorePassword": "changeit",
        "keyPassword": "changeit",
        "alias": "gridgrind-signing"
      }
    }
  }
}
```

`executionMode` is optional. Use it only when the request fits one of GridGrind's low-memory
contracts:

```json
{
  "executionMode": {
    "readMode": "EVENT_READ",
    "writeMode": "STREAMING_WRITE"
  }
}
```

- `EVENT_READ` supports only `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY` (`LIM-019`).
- `STREAMING_WRITE` requires `source.type: NEW` and supports only `ENSURE_SHEET`, `APPEND_ROW`,
  and `FORCE_FORMULA_RECALC_ON_OPEN` (`LIM-020`).

## Formula Environment

```json
{
  "formulaEnvironment": {
    "externalWorkbooks": [
      { "workbookName": "rates.xlsx", "path": "fixtures/rates.xlsx" }
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

## Coordinate Systems

| Field pattern | Convention |
|:--------------|:-----------|
| `address` | A1 cell address, e.g. `B3` |
| `range` | A1 rectangular range, e.g. `A1:C4` |
| `*RowIndex` | Zero-based row index, e.g. `0 = Excel row 1` |
| `*ColumnIndex` | Zero-based column index, e.g. `0 = Excel column A` |

Validation messages echo both forms inline, for example `firstRowIndex 5 (Excel row 6)` or
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
`'Sheet Name'!A1` syntax.

---

## Source

```json
{ "type": "NEW" }
{ "type": "EXISTING", "path": "path/to/file.xlsx" }
{ "type": "EXISTING", "path": "secured.xlsx", "security": { "password": "GridGrind-2026" } }
```

Relative `path` values resolve from the current working directory.

## Persistence

```json
{ "type": "SAVE_AS",   "path": "path/to/output.xlsx" }
{ "type": "OVERWRITE" }
{
  "type": "SAVE_AS",
  "path": "secured.xlsx",
  "security": {
    "encryption": { "password": "GridGrind-2026" }
  }
}
```

`SAVE_AS.path` writes a new file. `OVERWRITE` writes back to `source.path` and does not accept a
separate `path` field. `SAVE_AS` creates missing parent directories automatically. Signed
workbooks that are mutated require explicit `persistence.security.signature` re-signing before
they can be persisted. `GET_PACKAGE_SECURITY` is available only on the full-XSSF read path.

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

Sheet copy preserves non-drawing workbook-core structures such as tables, validations,
conditional formatting, comments, hyperlinks, local names, protection metadata, and print layout.
Drawing-family content such as pictures and charts remains outside the current copy contract.

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
  },
  "password": "Sheet-2026"
}
```

## CLEAR_SHEET_PROTECTION

```json
{ "type": "CLEAR_SHEET_PROTECTION", "sheetName": "Budget Review" }
```

## SET_WORKBOOK_PROTECTION

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

## CLEAR_WORKBOOK_PROTECTION

```json
{ "type": "CLEAR_WORKBOOK_PROTECTION" }
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

## INSERT_ROWS

```json
{
  "type": "INSERT_ROWS",
  "sheetName": "Sheet1",
  "rowIndex": 2,
  "rowCount": 3
}
```

Append-edge inserts on sparse sheets do not create a physical tail row until content or row
metadata exists there.

## DELETE_ROWS

```json
{
  "type": "DELETE_ROWS",
  "sheetName": "Sheet1",
  "rows": { "type": "BAND", "firstRowIndex": 4, "lastRowIndex": 6 }
}
```

Rejects deletes that would truncate a range-backed named range (`LIM-018`).

## SHIFT_ROWS

```json
{
  "type": "SHIFT_ROWS",
  "sheetName": "Sheet1",
  "rows": { "type": "BAND", "firstRowIndex": 1, "lastRowIndex": 3 },
  "delta": 2
}
```

Rejects shifts that would partially move or overwrite a range-backed named range outside the moved
band (`LIM-018`).

## INSERT_COLUMNS

```json
{
  "type": "INSERT_COLUMNS",
  "sheetName": "Sheet1",
  "columnIndex": 1,
  "columnCount": 2
}
```

Append-edge inserts on sparse sheets do not create a physical tail column until cells or explicit
column metadata exist there.

## DELETE_COLUMNS

```json
{
  "type": "DELETE_COLUMNS",
  "sheetName": "Sheet1",
  "columns": { "type": "BAND", "firstColumnIndex": 3, "lastColumnIndex": 4 }
}
```

Rejects deletes that would truncate a range-backed named range (`LIM-018`) and also rejects
formula-bearing workbooks (`LIM-017`).

## SHIFT_COLUMNS

```json
{
  "type": "SHIFT_COLUMNS",
  "sheetName": "Sheet1",
  "columns": { "type": "BAND", "firstColumnIndex": 0, "lastColumnIndex": 1 },
  "delta": -1
}
```

Rejects shifts that would partially move or overwrite a range-backed named range outside the moved
band (`LIM-018`) and also rejects formula-bearing workbooks (`LIM-017`).

## SET_ROW_VISIBILITY

```json
{
  "type": "SET_ROW_VISIBILITY",
  "sheetName": "Sheet1",
  "rows": { "type": "BAND", "firstRowIndex": 5, "lastRowIndex": 7 },
  "hidden": true
}
```

## SET_COLUMN_VISIBILITY

```json
{
  "type": "SET_COLUMN_VISIBILITY",
  "sheetName": "Sheet1",
  "columns": { "type": "BAND", "firstColumnIndex": 2, "lastColumnIndex": 3 },
  "hidden": false
}
```

## GROUP_ROWS

```json
{
  "type": "GROUP_ROWS",
  "sheetName": "Sheet1",
  "rows": { "type": "BAND", "firstRowIndex": 8, "lastRowIndex": 10 },
  "collapsed": true
}
```

## UNGROUP_ROWS

```json
{
  "type": "UNGROUP_ROWS",
  "sheetName": "Sheet1",
  "rows": { "type": "BAND", "firstRowIndex": 8, "lastRowIndex": 10 }
}
```

## GROUP_COLUMNS

```json
{
  "type": "GROUP_COLUMNS",
  "sheetName": "Sheet1",
  "columns": { "type": "BAND", "firstColumnIndex": 4, "lastColumnIndex": 6 },
  "collapsed": true
}
```

## UNGROUP_COLUMNS

```json
{
  "type": "UNGROUP_COLUMNS",
  "sheetName": "Sheet1",
  "columns": { "type": "BAND", "firstColumnIndex": 4, "lastColumnIndex": 6 }
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

## SET_SHEET_PRESENTATION

```json
{
  "type": "SET_SHEET_PRESENTATION",
  "sheetName": "Sheet1",
  "presentation": {
    "display": {
      "displayGridlines": false,
      "displayZeros": false,
      "displayRowColHeadings": true,
      "displayFormulas": false,
      "rightToLeft": false
    },
    "tabColor": { "rgb": "#0B6E4F" },
    "outlineSummary": { "rowSumsBelow": false, "rowSumsRight": true },
    "sheetDefaults": { "defaultColumnWidth": 14, "defaultRowHeightPoints": 19.0 },
    "ignoredErrors": [
      {
        "range": "B4:B18",
        "errorTypes": ["NUMBER_STORED_AS_TEXT"]
      }
    ]
  }
}
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
      "printGridlines": true,
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
{ "type": "SET_CELL", "sheetName": "Sheet1", "address": "H1", "value": { "type": "RICH_TEXT", "runs": [{ "text": "Q2 " }, { "text": "Budget", "font": { "bold": true, "fontColor": "#C00000" } }] } }
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
string in `path`. Relative `FILE` targets are analyzed against the workbook's persisted path when
one exists, so use absolute paths when you want cwd-independent health checks.

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

If `visible` is omitted, it defaults to `false`. When `runs` are present, their text must
concatenate exactly to `comment.text`.

## CLEAR_COMMENT

```json
{ "type": "CLEAR_COMMENT", "sheetName": "Sheet1", "address": "B2" }
```

## SET_PICTURE

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

Authored drawing mutations currently accept only `TWO_CELL` anchors with zero-based `from` and
`to` markers. `behavior` defaults to `MOVE_AND_RESIZE` when omitted.

## SET_SHAPE

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

`kind` is `SIMPLE_SHAPE` or `CONNECTOR`. `presetGeometryToken` and `text` are only for
`SIMPLE_SHAPE`, and `presetGeometryToken` defaults to `rect` when omitted.
Failed validation is non-mutating: unsupported preset geometry does not delete an existing object
or leave a partial new shape behind.

## SET_EMBEDDED_OBJECT

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

`previewImage` reuses the same `format` plus `base64Data` shape as `SET_PICTURE.picture.image`.

## SET_CHART

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

Formula-backed chart titles and series titles must resolve to one cell, either directly or through
one defined name that resolves to one cell. `categories.formula` and `values.formula` may still
target one contiguous range or one defined name that resolves to one contiguous range. Failed
validation is non-mutating: invalid chart payloads do not create partial charts or half-mutate an
existing supported chart.

Supported authored families are `BAR`, `LINE`, and `PIE`. Authored chart anchors currently accept
only `TWO_CELL`, and series formulas can point at contiguous ranges or defined names.

## SET_DRAWING_OBJECT_ANCHOR

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

## DELETE_DRAWING_OBJECT

```json
{ "type": "DELETE_DRAWING_OBJECT", "sheetName": "Ops", "objectName": "OpsConnector" }
```

## APPLY_STYLE

```json
{
  "type": "APPLY_STYLE",
  "sheetName": "Sheet1",
  "range": "A1:C1",
  "style": {
    "numberFormat": "#,##0.00",
    "alignment": {
      "wrapText": true,
      "horizontalAlignment": "CENTER",
      "verticalAlignment": "CENTER",
      "textRotation": 15,
      "indentation": 1
    },
    "font": {
      "bold": true,
      "italic": false,
      "fontName": "Aptos",
      "fontHeight": { "type": "POINTS", "points": 13 },
      "fontColor": "#1F4E78",
      "underline": true,
      "strikeout": false
    },
    "fill": {
      "pattern": "THIN_HORIZONTAL_BANDS",
      "foregroundColor": "#FFF2CC",
      "backgroundColor": "#FDE9D9"
    },
    "border": {
      "all": { "style": "THIN", "color": "#D6B656" },
      "right": { "style": "DOUBLE", "color": "#C55A11" }
    },
    "protection": {
      "locked": true,
      "hiddenFormula": false
    }
  }
}
```

```json
{
  "type": "APPLY_STYLE",
  "sheetName": "Sheet1",
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
      "bottom": { "style": "THIN", "colorIndexed": 8 }
    }
  }
}
```

`style` groups are `numberFormat`, `alignment`, `font`, `fill`, `border`, and `protection`.
`alignment.horizontalAlignment` values: `"LEFT"` `"CENTER"` `"RIGHT"` `"GENERAL"`
`alignment.verticalAlignment` values: `"TOP"` `"CENTER"` `"BOTTOM"`
`alignment.textRotation` uses explicit `0..180` degrees. `alignment.indentation` uses Excel's `0..250` cell-indent scale.
`font.fontHeight` accepts either `{ "type": "POINTS", "points": 11.5 }` or `{ "type": "TWIPS", "twips": 230 }`.
Color-bearing write fields accept RGB (`fontColor`, `fill.foregroundColor`, `fill.backgroundColor`, `border.*.color`), theme (`fontColorTheme`, `fill.foregroundColorTheme`, `fill.backgroundColorTheme`, `border.*.colorTheme`), or indexed (`fontColorIndexed`, `fill.foregroundColorIndexed`, `fill.backgroundColorIndexed`, `border.*.colorIndexed`) bases plus optional tint companions.
`fill.backgroundColor` and its theme/indexed variants are for patterned fills only; `SOLID` fills use foreground fields only.
`fill.gradient` authors gradient fills and is mutually exclusive with patterned fill fields.
`border.all` sets the default side style or color; explicit sides override it.
`border.*.color` requires a visible style on that side, either directly or via `border.all`.

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
    "ranges": ["K2:K20"],
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

Write support covers `FORMULA_RULE`, `CELL_VALUE_RULE`, `COLOR_SCALE_RULE`, `DATA_BAR_RULE`,
`ICON_SET_RULE`, and `TOP10_RULE`.

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
  "range": "A1:C20",
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
    "range": "A2:C20",
    "caseSensitive": false,
    "columnSort": false,
    "sortMethod": "",
    "conditions": [
      {
        "range": "C2:C20",
        "descending": true,
        "sortBy": ""
      }
    ]
  }
}
```

The range must include a nonblank header row and must not overlap any existing table range.
`criteria` defaults to `[]`; `sortState` is optional. `columnId` is zero-based within the
autofilter range.

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
    "hasAutofilter": false,
    "comment": "Dispatch tracker",
    "published": true,
    "insertRow": true,
    "insertRowShift": true,
    "headerRowCellStyle": "DispatchHeader",
    "dataCellStyle": "DispatchData",
    "totalsRowCellStyle": "DispatchTotals",
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

Table names are workbook-global. Header cells in the first row of the range must be nonblank and
unique case-insensitively. Overlapping different-name tables are rejected. Later value writes and
style patches that touch the table header row keep the table's persisted column metadata aligned
with the visible header cells. `showTotalsRow` is optional and defaults to `false`; when
`totalsRowFunction` is present in a column entry, input is case-insensitive and canonicalized to
Excel's lowercase stored token family.

## DELETE_TABLE

```json
{
  "type": "DELETE_TABLE",
  "name": "DispatchQueue",
  "sheetName": "Sheet1"
}
```

## SET_PIVOT_TABLE

```json
{
  "type": "SET_PIVOT_TABLE",
  "pivotTable": {
    "name": "SalesPivot",
    "sheetName": "Report",
    "source": {
      "type": "RANGE",
      "sheetName": "Sheet1",
      "range": "A1:D20"
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
{
  "type": "SET_PIVOT_TABLE",
  "pivotTable": {
    "name": "TablePivot",
    "sheetName": "TableReport",
    "source": { "type": "TABLE", "name": "DispatchQueue" },
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

Supported authored pivot sources are `RANGE`, `NAMED_RANGE`, and `TABLE`. `rowLabels`,
`columnLabels`, `reportFilters`, and `dataFields` must use disjoint source columns because POI
persists only one role per pivot field. When `reportFilters` is non-empty,
`anchor.topLeftAddress` must be on Excel row `3` or lower so the page-filter layout has room
above the rendered body. Supported `dataFields[*].function` values are `SUM`, `COUNT`,
`COUNT_NUMS`, `AVERAGE`, `MAX`, `MIN`, `PRODUCT`, `STD_DEV`, `STD_DEVP`, `VAR`, and `VARP`.

## DELETE_PIVOT_TABLE

```json
{
  "type": "DELETE_PIVOT_TABLE",
  "name": "SalesPivot",
  "sheetName": "Report"
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
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetRollup",
  "scope": { "type": "WORKBOOK" },
  "target": { "formula": "SUM(Budget!$B$2:$B$5)" }
}
```

Named-range targets are normalized to top-left:`bottom-right` order. For example, `"B4:A1"` is
accepted and stored as `"A1:B4"`. Formula-defined targets must set `formula` only.

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

## EVALUATE_FORMULA_CELLS

```json
{
  "type": "EVALUATE_FORMULA_CELLS",
  "cells": [
    { "sheetName": "Budget", "address": "D2" },
    { "sheetName": "Budget", "address": "E2" }
  ]
}
```

## CLEAR_FORMULA_CACHES

```json
{ "type": "CLEAR_FORMULA_CACHES" }
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

## GET_PACKAGE_SECURITY

```json
{ "type": "GET_PACKAGE_SECURITY", "requestId": "security" }
```

Returns factual OOXML package-encryption and OOXML package-signature state for the currently open
workbook. This read requires the full-XSSF path; `EVENT_READ` rejects it.

## GET_WORKBOOK_PROTECTION

```json
{ "type": "GET_WORKBOOK_PROTECTION", "requestId": "workbook-protection" }
```

Returns `structureLocked`, `windowsLocked`, `revisionLocked`, and whether workbook or revisions
password hashes are present.

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

String cells return `stringValue` and, when the stored cell contains authored rich text, an
optional ordered `richText` run list with effective per-run font facts.
Read-side style colors are structured objects with `rgb` plus optional `theme`, `indexed`, and
`tint`; gradient fills appear under `style.fill.gradient`.

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

Returned comments can include ordered rich-text `runs` plus an `anchor` with zero-based
`firstColumn`, `firstRow`, `lastColumn`, and `lastRow` bounds.

## GET_DRAWING_OBJECTS

```json
{ "type": "GET_DRAWING_OBJECTS", "requestId": "drawing-objects", "sheetName": "Ops" }
```

Returned entries are `PICTURE`, `CHART`, `SHAPE`, or `EMBEDDED_OBJECT`. `CHART` entries expose
`supported`, ordered `plotTypeTokens`, and title text. Read anchors can be `TWO_CELL`,
`ONE_CELL`, or `ABSOLUTE`.

## GET_CHARTS

```json
{ "type": "GET_CHARTS", "requestId": "charts", "sheetName": "Ops" }
```

Returned chart entries are `BAR`, `LINE`, `PIE`, or `UNSUPPORTED`. Supported simple charts include
ordered `axes` plus `series`; unsupported plot families preserve `plotTypeTokens` and `detail`.
Blank loaded chart titles normalize to `NONE`. Sparse literal caches surface missing positions as
empty strings. If a chart relation is gone but the graphic frame still exists, `GET_CHARTS` skips
the broken chart and `GET_DRAWING_OBJECTS` reports the surviving frame as read-only
`GRAPHIC_FRAME`.

## GET_DRAWING_OBJECT_PAYLOAD

```json
{
  "type": "GET_DRAWING_OBJECT_PAYLOAD",
  "requestId": "picture-payload",
  "sheetName": "Ops",
  "objectName": "OpsPicture"
}
{
  "type": "GET_DRAWING_OBJECT_PAYLOAD",
  "requestId": "embedded-payload",
  "sheetName": "Ops",
  "objectName": "OpsEmbed"
}
```

Payload extraction is only for named pictures and embedded objects. Non-binary drawing shapes such
as connectors and simple shapes are rejected.

## GET_SHEET_LAYOUT

```json
{ "type": "GET_SHEET_LAYOUT", "requestId": "layout", "sheetName": "Sheet1" }
```

Row and column entries report explicit size plus `hidden`, `outlineLevel`, and `collapsed`
where Excel exposes that state.

`layout.presentation` reports screen display flags, right-to-left layout, tab color,
outline-summary placement, sheet-default sizing, and ignored-error suppression.

## GET_PRINT_LAYOUT

```json
{ "type": "GET_PRINT_LAYOUT", "requestId": "print-layout", "sheetName": "Sheet1" }
```

`printLayout.setup` carries advanced page-setup facts such as margins, print-gridline output,
copies, first-page numbering, and explicit row or column breaks.

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

Entries are `SHEET` or `TABLE` and can include persisted `filterColumns` plus `sortState`.
Persisted sort-state metadata is reported exactly as stored, including blank raw ranges.

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
Malformed loaded rule metadata degrades to `UNSUPPORTED_RULE` instead of aborting the read.

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

Returned tables include persisted `columns` metadata plus advanced flags and style names such as
`comment`, `published`, `insertRow`, `insertRowShift`, `headerRowCellStyle`, `dataCellStyle`, and
`totalsRowCellStyle`.

## GET_PIVOT_TABLES

```json
{
  "type": "GET_PIVOT_TABLES",
  "requestId": "pivots",
  "selection": { "type": "ALL" }
}
{
  "type": "GET_PIVOT_TABLES",
  "requestId": "selected-pivots",
  "selection": {
    "type": "BY_NAMES",
    "names": ["SalesPivot", "NamedPivot"]
  }
}
```

Returned entries are `SUPPORTED` or `UNSUPPORTED`. Supported pivot reports include factual `source`,
stored `anchor`, `rowLabels`, `columnLabels`, `reportFilters`, `dataFields`, and
`valuesAxisOnColumns`. Unsupported or malformed loaded pivots are preserved as explicit
`UNSUPPORTED` entries with a human-readable `detail` instead of aborting the read.

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

## ANALYZE_PIVOT_TABLE_HEALTH

```json
{
  "type": "ANALYZE_PIVOT_TABLE_HEALTH",
  "requestId": "pivot-health",
  "selection": { "type": "ALL" }
}
{
  "type": "ANALYZE_PIVOT_TABLE_HEALTH",
  "requestId": "selected-pivot-health",
  "selection": {
    "type": "BY_NAMES",
    "names": ["SalesPivot", "NamedPivot"]
  }
}
```

Pivot health reports findings such as missing cache parts, missing workbook cache definitions,
broken sources, duplicate names, synthetic fallback names, or unsupported persisted detail.

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
conditional formatting, autofilter, table, pivot table, hyperlink, and named range. It is the
primary workbook-health check and pairs naturally with `persistence.type=NONE`.

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
  "type": "BY_NAMES",
  "names": ["SalesPivot", "NamedPivot"]
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
