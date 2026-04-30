---
afad: "4.0"
version: "0.62.0"
domain: DRAWING_STRUCTURED_INSPECTIONS
updated: "2026-05-01"
route:
  keywords: [gridgrind, inspections, get-drawing-objects, get-charts, get-tables, get-pivot-tables, get-sheet-schema]
  questions: ["how do i inspect drawings in gridgrind", "how do i inspect tables or pivots in gridgrind", "how do i inspect sheet layout or schema in gridgrind"]
---

# Drawing And Structured Inspection Reference

**Purpose**: Detailed reference for drawing, layout, validation, table, pivot, schema, and
structural inspection queries.
**Landing page**: [ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[ASSERTIONS.md](./ASSERTIONS.md), [WORKBOOK_AND_CELL_INSPECTIONS.md](./WORKBOOK_AND_CELL_INSPECTIONS.md),
and [ANALYSIS_QUERIES.md](./ANALYSIS_QUERIES.md)

### GET_DRAWING_OBJECTS

Returns factual drawing-object metadata for one sheet.

```json
{
  "stepId": "drawing-objects",
  "target": {
    "type": "DRAWING_OBJECT_ALL_ON_SHEET",
    "sheetName": "Ops"
  },
  "query": {
    "type": "GET_DRAWING_OBJECTS"
  }
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
- `SIGNATURE_LINE` with optional setup id, comment-permission flag, signer metadata, optional
  preview-image facts, and a factual `anchor`

Read-side anchors can be `TWO_CELL`, `ONE_CELL`, or `ABSOLUTE`. `TWO_CELL` markers expose
zero-based `columnIndex`, `rowIndex`, `dx`, and `dy`. `ONE_CELL` and `ABSOLUTE` anchors expose
their size fields in EMUs.

### GET_CHARTS

Returns factual chart metadata for one sheet. Supported authored plot families and multi-plot
combinations built from them are modeled authoritatively. Unsupported loaded plot families still
surface as explicit `UNSUPPORTED` plot entries with preserved detail inside the returned chart.

```json
{
  "stepId": "charts",
  "target": {
    "type": "CHART_ALL_ON_SHEET",
    "sheetName": "Ops"
  },
  "query": {
    "type": "GET_CHARTS"
  }
}
```

Response shape: `{ "charts": [ ... ] }`.

Returned chart entries include chart-level `name`, `anchor`, `title`, `legend`,
`displayBlanksAs`, `plotOnlyVisibleCells`, and ordered `plots`.

Returned `plots` entries are one of:

- `AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`, `LINE_3D`, `PIE`, `PIE_3D`, `RADAR`,
  `SCATTER`, `SURFACE`, or `SURFACE_3D` with factual plot-specific properties plus ordered
  `series`
- `UNSUPPORTED` with preserved `plotTypeTokens` and human-readable `detail`

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
  "stepId": "picture-payload",
  "target": {
    "type": "DRAWING_OBJECT_BY_NAME",
    "sheetName": "Ops",
    "objectName": "OpsPicture"
  },
  "query": {
    "type": "GET_DRAWING_OBJECT_PAYLOAD"
  }
}
```

```json
{
  "stepId": "embedded-payload",
  "target": {
    "type": "DRAWING_OBJECT_BY_NAME",
    "sheetName": "Ops",
    "objectName": "OpsEmbed"
  },
  "query": {
    "type": "GET_DRAWING_OBJECT_PAYLOAD"
  }
}
```

Response shape: `{ "payload": { ... } }`.

Returned payload entries are:

- `PICTURE` with `format`, `contentType`, `fileName`, `sha256`, `base64Data`, and optional
  `description`
- `EMBEDDED_OBJECT` with `packagingKind`, `contentType`, optional `fileName`, `sha256`,
  `base64Data`, and optional `label` or `command`

Named non-binary drawing objects such as signature lines, connectors, and simple shapes are
rejected because they do not own an extractable binary payload.

### GET_SHEET_LAYOUT

Returns pane state, effective zoom, sheet-presentation state, and per-row or per-column layout
facts for one sheet. Row and column entries include explicit size plus `hidden`, `outlineLevel`,
and `collapsed` where Excel exposes that state.

```json
{
  "stepId": "layout",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "query": {
    "type": "GET_SHEET_LAYOUT"
  }
}
```

The returned `layout.pane` is one of:

- `NONE`
- `FROZEN` with `splitColumn`, `splitRow`, `leftmostColumn`, and `topRow`
- `SPLIT` with `xSplitPosition`, `ySplitPosition`, `leftmostColumn`, `topRow`, and `activePane`

The returned `layout.presentation` object reports:

- `display`: `displayGridlines`, `displayZeros`, `displayRowColHeadings`, `displayFormulas`, and
  `rightToLeft`
- `tabColor`: omitted when the sheet has no tab color, or a structured color report when one is set
- `outlineSummary`: `rowSumsBelow` and `rowSumsRight`
- `sheetDefaults`: `defaultColumnWidth` and `defaultRowHeightPoints`
- `ignoredErrors`: factual ignored-error blocks grouped by range

`GET_SHEET_LAYOUT` is factual rather than revalidated. If a workbook already contains malformed but
positive persisted explicit row heights, explicit column widths, or default row-height values,
those values are returned as stored instead of being clamped to the normal mutation-time limits.

### GET_PRINT_LAYOUT

Returns the supported print-layout state for one sheet.

```json
{
  "stepId": "print-layout",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "query": {
    "type": "GET_PRINT_LAYOUT"
  }
}
```

The returned `printLayout.setup` object carries advanced page-setup facts: `margins`,
`printGridlines`, `horizontallyCentered`, `verticallyCentered`, `paperSize`, `draft`,
`blackAndWhite`, `copies`, `useFirstPageNumber`, `firstPageNumber`, and explicit `rowBreaks`
plus `columnBreaks`.

### GET_DATA_VALIDATIONS

Returns factual data-validation structures for one sheet. Each returned entry is one of:

- `SUPPORTED`: a fully modeled validation definition plus its normalized stored ranges
- `UNSUPPORTED`: a present workbook rule GridGrind can detect but cannot expose as a supported
  validation definition; the entry includes `kind` and `detail` so callers can distinguish
  unsupported families from invalid workbook structures that Apache POI still exposes

```json
{
  "stepId": "data-validations",
  "target": {
    "type": "RANGE_ALL_ON_SHEET",
    "sheetName": "Inventory"
  },
  "query": {
    "type": "GET_DATA_VALIDATIONS"
  }
}
```

```json
{
  "stepId": "selected-data-validations",
  "target": {
    "type": "RANGE_BY_RANGES",
    "sheetName": "Inventory",
    "ranges": [
      "B2:B200",
      "C2:C200"
    ]
  },
  "query": {
    "type": "GET_DATA_VALIDATIONS"
  }
}
```

Range-selector payloads use:

```json
{
  "type": "RANGE_ALL_ON_SHEET",
  "sheetName": "Inventory"
}
{
  "type": "RANGE_BY_RANGES",
  "sheetName": "Inventory",
  "ranges": [
    "B2:B200",
    "C2:C200"
  ]
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
  "stepId": "conditional-formatting",
  "target": {
    "type": "RANGE_ALL_ON_SHEET",
    "sheetName": "Inventory"
  },
  "query": {
    "type": "GET_CONDITIONAL_FORMATTING"
  }
}
```

```json
{
  "stepId": "selected-conditional-formatting",
  "target": {
    "type": "RANGE_BY_RANGES",
    "sheetName": "Inventory",
    "ranges": [
      "A2:D200",
      "F2:F20"
    ]
  },
  "query": {
    "type": "GET_CONDITIONAL_FORMATTING"
  }
}
```

### GET_AUTOFILTERS

Returns factual autofilter metadata for one sheet. The result may include:

- `SHEET`: one sheet-owned autofilter stored directly on the worksheet
- `TABLE`: one table-owned autofilter stored on a table definition, including `tableName`

```json
{
  "stepId": "autofilters",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "query": {
    "type": "GET_AUTOFILTERS"
  }
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
  "stepId": "tables",
  "target": {
    "type": "TABLE_ALL"
  },
  "query": {
    "type": "GET_TABLES"
  }
}
```

```json
{
  "stepId": "selected-tables",
  "target": {
    "type": "TABLE_BY_NAMES",
    "names": [
      "InventoryTable",
      "Trips"
    ]
  },
  "query": {
    "type": "GET_TABLES"
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

Table-selector payloads use:

```json
{
  "type": "TABLE_ALL"
}
{
  "type": "TABLE_BY_NAMES",
  "names": [
    "InventoryTable",
    "Trips"
  ]
}
```

### GET_PIVOT_TABLES

Returns factual pivot-table metadata selected by workbook-global pivot-table name or all pivots.
Supported pivots surface source, stored anchor, row or column labels, report filters, data fields,
and values-axis placement. Unsupported or malformed loaded pivots are returned explicitly with
preserved detail instead of causing read failure.

```json
{
  "stepId": "pivots",
  "target": {
    "type": "PIVOT_TABLE_ALL"
  },
  "query": {
    "type": "GET_PIVOT_TABLES"
  }
}
```

```json
{
  "stepId": "selected-pivots",
  "target": {
    "type": "PIVOT_TABLE_BY_NAMES",
    "names": [
      "SalesPivot",
      "NamedPivot"
    ]
  },
  "query": {
    "type": "GET_PIVOT_TABLES"
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
  "stepId": "formula-surface",
  "target": {
    "type": "SHEET_ALL"
  },
  "query": {
    "type": "GET_FORMULA_SURFACE"
  }
}
```

```json
{
  "stepId": "selected-formula-surface",
  "target": {
    "type": "SHEET_BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "GET_FORMULA_SURFACE"
  }
}
```

Response shape: `{ "analysis": { "totalFormulaCellCount": ..., "sheets": [ { "sheetName": ...,`
`"formulaCellCount": ..., "distinctFormulaCount": ..., "formulas": [ { "formula": ...,`
`"occurrenceCount": ..., "addresses": [...] } ] } ] } }`.

Sheet-selector payloads use:

```json
{
  "type": "SHEET_ALL"
}
{
  "type": "SHEET_BY_NAMES",
  "names": [
    "Inventory",
    "Summary"
  ]
}
```

### GET_SHEET_SCHEMA

Infers a simple schema from one rectangular sheet window. The first row of the window is treated
as a header row; `dataRowCount` in the response is `rowCount - 1`. When the header row is
entirely blank, `dataRowCount` is `0`.

`rowCount * columnCount` must not exceed 250,000. Requests that exceed this limit are rejected
with `INVALID_REQUEST`.

`dominantType` is omitted per column when all data cells are blank, or when two or more types tie
for the highest count. Formula cells contribute their evaluated result type (e.g. `NUMBER`,
`STRING`) to `observedTypes` and `dominantType`, not `FORMULA`.

```json
{
  "stepId": "schema",
  "target": {
    "type": "RANGE_RECTANGULAR_WINDOW",
    "sheetName": "Inventory",
    "topLeftAddress": "A1",
    "rowCount": 5,
    "columnCount": 3
  },
  "query": {
    "type": "GET_SHEET_SCHEMA"
  }
}
```

### GET_NAMED_RANGE_SURFACE

Summarizes the scope and backing kind of selected named ranges.

```json
{
  "stepId": "named-range-surface",
  "target": {
    "type": "NAMED_RANGE_ALL"
  },
  "query": {
    "type": "GET_NAMED_RANGE_SURFACE"
  }
}
```

```json
{
  "stepId": "selected-named-range-surface",
  "target": {
    "type": "NAMED_RANGE_ANY_OF",
    "selectors": [
      {
        "type": "NAMED_RANGE_WORKBOOK_SCOPE",
        "name": "BudgetTotal"
      },
      {
        "type": "NAMED_RANGE_SHEET_SCOPE",
        "sheetName": "Budget",
        "name": "BudgetTable"
      }
    ]
  },
  "query": {
    "type": "GET_NAMED_RANGE_SURFACE"
  }
}
```

Response shape: `{ "analysis": { "workbookScopedCount": ..., "sheetScopedCount": ...,`
`"rangeBackedCount": ..., "formulaBackedCount": ..., "namedRanges": [ { "name": ..., "scope":`
`..., "refersToFormula": ..., "kind": ... } ] } }`.
