---
afad: "3.5"
version: "0.61.0"
domain: STRUCTURED_DATA_MUTATIONS
updated: "2026-04-25"
route:
  keywords: [gridgrind, autofilter, table, pivot table, append-row, auto-size, execution calculation, named range]
  questions: ["how do i manage tables in gridgrind", "how do i set pivot tables in gridgrind", "how do i use append row or calculation policies in gridgrind"]
---

# Structured Data Mutation Reference

**Purpose**: Detailed mutation reference for autofilters, tables, pivot tables, append-only
streaming writes, calculation policy wiring, and named ranges.
**Landing page**: [STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[STYLE_AND_VALIDATION_MUTATIONS.md](./STYLE_AND_VALIDATION_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

### SET_AUTOFILTER

Create or replace one sheet-level autofilter range. The range must be rectangular, include a
nonblank header row, and must not overlap any existing table range on the same sheet. Optional
`criteria` author persisted filter-column rules; optional `sortState` authors persisted sort-state
metadata on the same autofilter.

```json
{
  "stepId": "set-autofilter",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "A1:C200"
  },
  "action": {
    "type": "SET_AUTOFILTER",
    "criteria": [
      {
        "columnId": 2,
        "showButton": true,
        "criterion": {
          "type": "VALUES",
          "values": [
            "Queued",
            "Done"
          ],
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
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
{
  "stepId": "clear-autofilter",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "CLEAR_AUTOFILTER"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAME` selector for the existing target sheet. |

---

### SET_TABLE

Create or replace one workbook-global table definition. The first row of `table.range` supplies
the table's header cells, which must be nonblank and unique case-insensitively. Same-name tables
on the same sheet are replaced. Same-name tables on a different sheet are rejected. Overlapping
different-name tables are rejected. If the new table overlaps a sheet-level autofilter, the
sheet-level filter is cleared so the table-owned autofilter becomes authoritative on that range.
Later value writes and style patches that touch the table header row keep the persisted
table-column metadata converged with the visible header cells.
The step target must be `TableSelector.BY_NAME_ON_SHEET`, and the selector must match
`table.name` plus `table.sheetName`.

```json
{
  "type": "SET_TABLE",
  "table": {
    "name": "InventoryTable",
    "sheetName": "Inventory",
    "range": "A1:C200",
    "style": {
      "type": "NONE"
    }
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
    "comment": {
      "type": "INLINE",
      "text": "Inventory tracker"
    },
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
      {
        "columnIndex": 0,
        "totalsRowLabel": "Total"
      },
      {
        "columnIndex": 1,
        "totalsRowFunction": "SUM"
      },
      {
        "columnIndex": 2,
        "uniqueName": "status-unique"
      }
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
The step target must be `PivotTableSelector.BY_NAME_ON_SHEET`, and the selector must match
`pivotTable.name` plus `pivotTable.sheetName`.

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
    "anchor": {
      "topLeftAddress": "C5"
    },
    "rowLabels": [
      "Region"
    ],
    "columnLabels": [
      "Stage"
    ],
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
    "source": {
      "type": "NAMED_RANGE",
      "name": "PivotSource"
    },
    "anchor": {
      "topLeftAddress": "A3"
    },
    "rowLabels": [
      "Region"
    ],
    "columnLabels": [],
    "reportFilters": [
      "Owner"
    ],
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
    "source": {
      "type": "TABLE",
      "name": "InventoryTable"
    },
    "anchor": {
      "topLeftAddress": "F4"
    },
    "rowLabels": [
      "Stage"
    ],
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
  "stepId": "append-row",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "APPEND_ROW",
    "values": [
      {
        "type": "TEXT",
        "source": {
          "type": "INLINE",
          "text": "Guatemala Antigua"
        }
      },
      {
        "type": "NUMBER",
        "number": 100
      },
      {
        "type": "NUMBER",
        "number": 9.2
      }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAME` selector for the target sheet. |
| `values` | Yes | Row of typed cell values. |

Rows that contain only style, comment, or hyperlink metadata are ignored when locating the append
position.

In `execution.mode.writeMode=STREAMING_WRITE`, `APPEND_ROW` is valid only after an
`ENSURE_SHEET` mutation has already created or selected the target sheet. The streaming contract
requires at least one `ENSURE_SHEET` step before append/assert/read work begins.

---

### AUTO_SIZE_COLUMNS

Resize columns to fit their content. Applies to all columns with data on the sheet.

```json
{
  "stepId": "auto-size-columns",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "AUTO_SIZE_COLUMNS"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAME` selector for the target sheet. |

Note: `AUTO_SIZE_COLUMNS` and `SET_COLUMN_WIDTH` can be combined in the same request. Since
operations run in order, whichever appears later wins.

GridGrind uses deterministic content-based sizing rather than host font metrics, so Docker,
headless, and local runs produce the same widths.

---

### execution.calculation

Formula evaluation, cache clearing, and workbook-open recalc are now request-level execution
policy, not mutation actions. One calculation policy applies to the whole request and runs after
mutations, before any downstream assertion or inspection observes formula state.

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "EVALUATE_ALL"
      },
      "markRecalculateOnOpen": true
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `strategy` | No | `DO_NOT_CALCULATE`, `EVALUATE_ALL`, `EVALUATE_TARGETS`, or `CLEAR_CACHES_ONLY`. Defaults to `DO_NOT_CALCULATE`. |
| `markRecalculateOnOpen` | No | Persist Excel's workbook-level recalc-on-open flag. Defaults to `false`. |

`EVALUATE_ALL` refreshes every formula cell reachable in the workbook:

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "EVALUATE_ALL"
      }
    }
  }
}
```

`EVALUATE_TARGETS` refreshes one explicit formula-cell set only:

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "EVALUATE_TARGETS",
        "cells": [
          {
            "sheetName": "Budget",
            "address": "D2"
          },
          {
            "sheetName": "Budget",
            "address": "E2"
          }
        ]
      }
    }
  }
}
```

Every `cells[*]` entry must point at an existing formula cell. A missing physical cell can surface
`CELL_NOT_FOUND`; an existing non-formula cell is rejected as `INVALID_REQUEST`.

`CLEAR_CACHES_ONLY` strips persisted formula caches without attempting server-side evaluation:

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "CLEAR_CACHES_ONLY"
      }
    }
  }
}
```

`markRecalculateOnOpen=true` can be paired with any strategy, including the default
`DO_NOT_CALCULATE`, when Excel-compatible clients should refresh formulas later instead of the
server evaluating them immediately.

`EVENT_READ` requires `strategy=DO_NOT_CALCULATE` with `markRecalculateOnOpen=false`.
`STREAMING_WRITE` requires `strategy=DO_NOT_CALCULATE` and allows only `markRecalculateOnOpen=true`
as its calculation-side change.

---

### SET_NAMED_RANGE

Create or replace one named range in workbook scope or sheet scope. Targets are explicit
sheet-qualified cells or rectangular ranges, or formula-defined name targets.

```json
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetTotal",
  "scope": {
    "type": "WORKBOOK"
  },
  "target": {
    "sheetName": "Budget",
    "range": "B4"
  }
}
```

```json
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetTable",
  "scope": {
    "type": "SHEET",
    "sheetName": "Budget"
  },
  "target": {
    "sheetName": "Budget",
    "range": "A1:B4"
  }
}
```

```json
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetRollup",
  "scope": {
    "type": "WORKBOOK"
  },
  "target": {
    "formula": "SUM(Budget!$B$2:$B$5)"
  }
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
  "scope": {
    "type": "WORKBOOK"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Defined-name identifier to delete. |
| `scope` | Yes | Workbook or sheet scope of the exact name to delete. |

---
