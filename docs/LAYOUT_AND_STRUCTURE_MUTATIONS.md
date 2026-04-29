---
afad: "3.5"
version: "0.61.0"
domain: LAYOUT_STRUCTURE_MUTATIONS
updated: "2026-04-25"
route:
  keywords: [gridgrind, layout mutations, merge-cells, rows, columns, pane, zoom, presentation, print-layout]
  questions: ["how do i change workbook layout in gridgrind", "how do i insert or delete rows in gridgrind", "how do i set panes or print layout in gridgrind"]
---

# Layout And Structure Mutation Reference

**Purpose**: Detailed mutation reference for merges, row and column structure, visibility,
panes, zoom, presentation, and print-layout state.
**Landing page**: [WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[WORKBOOK_AND_SHEET_MUTATIONS.md](./WORKBOOK_AND_SHEET_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

### MERGE_CELLS

Merge a rectangular A1-style range into one displayed region. The range must span at least two
cells. Repeating the same merge is a no-op, but overlapping a different merged region fails.

```json
{
  "stepId": "merge-cells",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "A1:C1"
  },
  "action": {
    "type": "MERGE_CELLS"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `range` | Yes | A1-notation rectangular range to merge. Must span at least two cells. |

---

### UNMERGE_CELLS

Remove the merged region whose coordinates exactly match `range`. GridGrind does not partially
unmerge intersecting regions; the range must identify the merged region exactly.

```json
{
  "stepId": "unmerge-cells",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "A1:C1"
  },
  "action": {
    "type": "UNMERGE_CELLS"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
| `target` | Yes | Selector payload for the target workbook location. |
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
| `target` | Yes | Selector payload for the target workbook location. |
| `firstRowIndex` | Yes | Zero-based first row index. |
| `lastRowIndex` | Yes | Zero-based last row index. Must be greater than or equal to `firstRowIndex`. |
| `heightPoints` | Yes | Positive Excel row height in points. Must be finite and > 0 and ≤ 409.0 (Excel row height limit). |

---

### INSERT_ROWS

Insert one or more blank rows before `rowIndex`. `rowIndex` is zero-based and may point at the
current append position (`last existing row + 1`) but not beyond it. On sparse sheets, inserting
at the append edge does not materialize a new physical tail row until content or row metadata
exists there.

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
| `target` | Yes | Selector payload for the target workbook location. |
| `rowIndex` | Yes | Zero-based insertion point. Must be within the current row bounds or exactly one past the last existing row. |
| `rowCount` | Yes | Positive number of blank rows to insert. |

GridGrind preserves and retargets existing data validations across the insert so the moved cells
keep their authored validation coverage. It still rejects the edit when Apache POI would leave a
moved table or sheet-owned autofilter stale (`LIM-016`).

---

### DELETE_ROWS

Delete one inclusive zero-based row band. Rows beneath the deleted band shift upward.

```json
{
  "type": "DELETE_ROWS",
  "sheetName": "Inventory",
  "rows": {
    "type": "BAND",
    "firstRowIndex": 4,
    "lastRowIndex": 6
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "rows": {
    "type": "BAND",
    "firstRowIndex": 1,
    "lastRowIndex": 3
  },
  "delta": 2
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `rows` | Yes | One inclusive row-band payload. |
| `delta` | Yes | Signed non-zero row offset. |

GridGrind rejects the edit when it would move a table, sheet-owned autofilter, or data
validation (`LIM-016`), and it also rejects shifts that would partially move or overwrite a
range-backed named range outside the moved band (`LIM-018`).

---

### INSERT_COLUMNS

Insert one or more blank columns before `columnIndex`. `columnIndex` is zero-based and may point
at the current append position (`last existing column + 1`) but not beyond it. On sparse sheets,
inserting at the append edge does not materialize a new physical tail column until cells or
explicit column metadata exist there.

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
| `target` | Yes | Selector payload for the target workbook location. |
| `columnIndex` | Yes | Zero-based insertion point. Must be within the current column bounds or exactly one past the last existing column. |
| `columnCount` | Yes | Positive number of blank columns to insert. |

GridGrind preserves and retargets existing data validations across the insert so the moved cells
keep their authored validation coverage. It still rejects the edit when Apache POI would leave a
moved table or sheet-owned autofilter stale (`LIM-016`), and it also rejects the edit when the
workbook contains any formula cells or formula-defined named ranges because Apache POI leaves some
column references stale after column moves (`LIM-017`).

---

### DELETE_COLUMNS

Delete one inclusive zero-based column band. Columns to the right of the deleted band shift left.

```json
{
  "type": "DELETE_COLUMNS",
  "sheetName": "Inventory",
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 3,
    "lastColumnIndex": 4
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 0,
    "lastColumnIndex": 1
  },
  "delta": -1
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "rows": {
    "type": "BAND",
    "firstRowIndex": 5,
    "lastRowIndex": 7
  },
  "hidden": true
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `rows` | Yes | One inclusive row-band payload. |
| `hidden` | Yes | `true` hides the rows; `false` unhides them. |

---

### SET_COLUMN_VISIBILITY

Hide or unhide one inclusive zero-based column band.

```json
{
  "type": "SET_COLUMN_VISIBILITY",
  "sheetName": "Inventory",
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 2,
    "lastColumnIndex": 3
  },
  "hidden": false
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "rows": {
    "type": "BAND",
    "firstRowIndex": 8,
    "lastRowIndex": 10
  },
  "collapsed": true
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "rows": {
    "type": "BAND",
    "firstRowIndex": 8,
    "lastRowIndex": 10
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `rows` | Yes | One inclusive row-band payload. |

---

### GROUP_COLUMNS

Apply one outline group to one inclusive zero-based column band. When `collapsed` is omitted it
defaults to `false`.

```json
{
  "type": "GROUP_COLUMNS",
  "sheetName": "Inventory",
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 4,
    "lastColumnIndex": 6
  },
  "collapsed": true
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 4,
    "lastColumnIndex": 6
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `columns` | Yes | One inclusive column-band payload. |

---

### SET_SHEET_PANE

Apply one explicit pane state to a sheet. Use `NONE` to clear panes, `FROZEN` for freeze-pane
behavior, or `SPLIT` for Excel split panes.

```json
{
  "stepId": "set-sheet-pane",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_SHEET_PANE",
    "pane": {
      "type": "FROZEN",
      "splitColumn": 1,
      "splitRow": 1,
      "leftmostColumn": 1,
      "topRow": 1
    }
  }
}
```

```json
{
  "stepId": "set-sheet-pane",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_SHEET_PANE",
    "pane": {
      "type": "SPLIT",
      "xSplitPosition": 2400,
      "ySplitPosition": 1800,
      "leftmostColumn": 2,
      "topRow": 3,
      "activePane": "LOWER_RIGHT"
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "stepId": "set-sheet-zoom",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_SHEET_ZOOM",
    "zoomPercent": 125
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `zoomPercent` | Yes | Integer zoom percentage. Must be between `10` and `400` inclusive. |

### SET_SHEET_PRESENTATION

Apply one authoritative supported sheet-presentation state to a sheet. Omitted nested fields
normalize to defaults or clear state.

```json
{
  "stepId": "set-sheet-presentation",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_SHEET_PRESENTATION",
    "presentation": {
      "display": {
        "displayGridlines": false,
        "displayZeros": false,
        "displayRowColHeadings": true,
        "displayFormulas": false,
        "rightToLeft": false
      },
      "tabColor": {
        "rgb": "#0B6E4F"
      },
      "outlineSummary": {
        "rowSumsBelow": false,
        "rowSumsRight": true
      },
      "sheetDefaults": {
        "defaultColumnWidth": 14,
        "defaultRowHeightPoints": 19.0
      },
      "ignoredErrors": [
        {
          "range": "B4:B18",
          "errorTypes": [
            "NUMBER_STORED_AS_TEXT"
          ]
        }
      ]
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `presentation` | Yes | Authoritative sheet-presentation payload. |

`presentation` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `display` | No | Screen-view flags. Each nested field is optional and defaults to Excel's worksheet defaults. |
| `tabColor` | No | Tab color using the normal `ColorInput` contract. Omit to clear the tab color. |
| `outlineSummary` | No | `rowSumsBelow` and `rowSumsRight`. Defaults to `true` for both. |
| `sheetDefaults` | No | Sheet-wide fallback sizing for rows and columns without explicit overrides. |
| `ignoredErrors` | No | Authoritative list of ignored-error blocks. Omit or pass `[]` to clear all ignored-error suppression. |

Nested constraints:

- `display` supports `displayGridlines`, `displayZeros`, `displayRowColHeadings`,
  `displayFormulas`, and `rightToLeft`.
- `sheetDefaults.defaultColumnWidth` must be a whole number greater than `0` and less than or equal to `255`.
- `sheetDefaults.defaultRowHeightPoints` must be finite, greater than `0`, and less than or equal to `409.0`.
- Each `ignoredErrors` entry requires one A1-style rectangular `range` plus one or more distinct
  `errorTypes`.
- `ignoredErrors` ranges must be unique within the request.

Those ceilings apply to authored mutations. `GET_SHEET_LAYOUT` remains factual on reopen, so
malformed positive persisted explicit row or column sizes, plus malformed positive persisted
default row height, are reported as stored instead of being clamped back to authoring limits.

### SET_PRINT_LAYOUT

Apply one authoritative supported print-layout state to a sheet. Omitted nested fields normalize
to default or cleared state.

```json
{
  "stepId": "set-print-layout",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_PRINT_LAYOUT",
    "printLayout": {
      "printArea": {
        "type": "RANGE",
        "range": "A1:F40"
      },
      "orientation": "LANDSCAPE",
      "scaling": {
        "type": "FIT",
        "widthPages": 1,
        "heightPages": 0
      },
      "repeatingRows": {
        "type": "BAND",
        "firstRowIndex": 0,
        "lastRowIndex": 1
      },
      "repeatingColumns": {
        "type": "BAND",
        "firstColumnIndex": 0,
        "lastColumnIndex": 0
      },
      "header": {
        "left": {
          "type": "INLINE",
          "text": "Inventory"
        },
        "center": {
          "type": "INLINE",
          "text": "Q2"
        },
        "right": {
          "type": "INLINE",
          "text": "Internal"
        }
      },
      "footer": {
        "left": {
          "type": "INLINE",
          "text": ""
        },
        "center": {
          "type": "INLINE",
          "text": "Prepared by GridGrind"
        },
        "right": {
          "type": "INLINE",
          "text": "Page 1"
        }
      },
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
        "rowBreaks": [
          6
        ],
        "columnBreaks": [
          3
        ]
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
| `setup` | No | Advanced page-setup payload. Defaults to the supported Excel baseline for margins, print-gridline output, centering, copies, and page numbering. |

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
  "stepId": "clear-print-layout",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "CLEAR_PRINT_LAYOUT"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

---
