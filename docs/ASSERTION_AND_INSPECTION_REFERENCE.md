---
afad: "3.5"
version: "0.53.0"
domain: ASSERTION_INSPECTION_REFERENCE
updated: "2026-04-22"
route:
  keywords: [gridgrind, assertions, inspections, query, get-cells, get-sheet-layout, analysis, workbook-health, formula-health, named-range-health, charts, drawing-objects]
  questions: ["what assertions does gridgrind support", "what inspection queries does gridgrind support", "how do I read cells in gridgrind", "how do I inspect drawing objects in gridgrind", "how do I run workbook health checks in gridgrind"]
---

# Assertion And Inspection Reference

**Purpose**: Long-form reference for assertion steps, factual inspection queries, and finding-
bearing analysis queries.
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md),
[CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md), and
[STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md)

## Assertions

Assertions are ordered, explicit verification steps. They run in the same `steps[]` list as
mutations and inspections, and they can appear anywhere the plan needs them. Every assertion must
include a caller-defined `stepId`.

Successful responses echo passed assertion steps back through the ordered `assertions[]` array:

```json
{
  "status": "SUCCESS",
  "protocolVersion": "V1",
  "journal": {
    "planId": "assert-budget",
    "level": "NORMAL"
  },
  "persistence": {
    "type": "NONE"
  },
  "warnings": [],
  "assertions": [
    {
      "stepId": "assert-title",
      "assertionType": "EXPECT_CELL_VALUE"
    }
  ],
  "inspections": []
}
```

Failed assertions stop the workflow with `ASSERTION_FAILED` and attach a structured
`problem.assertionFailure` payload describing the failed assertion plus the observed factual read
results that caused the mismatch.

`EXPECT_PRESENT` and `EXPECT_ABSENT` are selector-count assertions, not strict read lookups. If an
exact named range, chart, table, or pivot-table selector matches nothing, the assertion observes
zero entities and then passes or fails from that count; GridGrind does not surface
selector-specific `*_NOT_FOUND` problems for these assertion families.

Assertion families:

| Assertion `type` | Valid target families | Purpose |
|:-----------------|:----------------------|:--------|
| `EXPECT_PRESENT` | `NamedRangeSelector`, `TableSelector`, `PivotTableSelector`, `ChartSelector` | Require at least one matching workbook entity; selector misses count as zero observed entities. |
| `EXPECT_ABSENT` | `NamedRangeSelector`, `TableSelector`, `PivotTableSelector`, `ChartSelector` | Require zero matching workbook entities; selector misses count as zero observed entities. |
| `EXPECT_CELL_VALUE` | `CellSelector` | Require exact effective cell values. |
| `EXPECT_DISPLAY_VALUE` | `CellSelector` | Require exact formatted display strings. |
| `EXPECT_FORMULA_TEXT` | `CellSelector` | Require exact stored formula text. |
| `EXPECT_CELL_STYLE` | `CellSelector` | Require exact cell-style reports. |
| `EXPECT_WORKBOOK_PROTECTION` | `WorkbookSelector` | Require exact workbook-protection facts. |
| `EXPECT_SHEET_STRUCTURE` | `SheetSelector.ByName` | Require exact one-sheet structural summary facts. |
| `EXPECT_NAMED_RANGE_FACTS` | `NamedRangeSelector` | Require exact named-range reports. |
| `EXPECT_TABLE_FACTS` | `TableSelector` | Require exact table reports. |
| `EXPECT_PIVOT_TABLE_FACTS` | `PivotTableSelector` | Require exact pivot-table reports. |
| `EXPECT_CHART_FACTS` | `ChartSelector` | Require exact chart reports. |
| `EXPECT_ANALYSIS_MAX_SEVERITY` | Matches the supplied analysis `query` target family | Require a maximum severity ceiling for one analysis query. |
| `EXPECT_ANALYSIS_FINDING_PRESENT` | Matches the supplied analysis `query` target family | Require at least one matching finding from one analysis query. |
| `EXPECT_ANALYSIS_FINDING_ABSENT` | Matches the supplied analysis `query` target family | Require no matching findings from one analysis query. |
| `ALL_OF` | Any selector family shared by all nested assertions | Require every nested assertion to pass against the same target. |
| `ANY_OF` | Any selector family shared by all nested assertions | Require at least one nested assertion to pass against the same target. |
| `NOT` | Same selector family as the nested assertion | Invert one nested assertion. |

Common assertion step shapes:

```json
{
  "stepId": "assert-title",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Budget",
    "address": "A1"
  },
  "assertion": {
    "type": "EXPECT_CELL_VALUE",
    "expectedValue": {
      "type": "TEXT",
      "text": "Quarterly Budget"
    }
  }
}
```

```json
{
  "stepId": "assert-formula-health",
  "target": {
    "type": "BY_NAME",
    "sheetName": "Budget"
  },
  "assertion": {
    "type": "EXPECT_ANALYSIS_MAX_SEVERITY",
    "query": {
      "type": "ANALYZE_FORMULA_HEALTH"
    },
    "maximumSeverity": "INFO"
  }
}
```

```json
{
  "stepId": "assert-workbook-links",
  "target": {
    "type": "CURRENT"
  },
  "assertion": {
    "type": "EXPECT_ANALYSIS_FINDING_ABSENT",
    "query": {
      "type": "ANALYZE_WORKBOOK_FINDINGS"
    },
    "code": "HYPERLINK_MALFORMED_TARGET"
  }
}
```

```json
{
  "stepId": "assert-total-state",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Budget",
    "address": "B2"
  },
  "assertion": {
    "type": "ALL_OF",
    "assertions": [
      {
        "type": "EXPECT_CELL_VALUE",
        "expectedValue": {
          "type": "NUMBER",
          "number": 1200
        }
      },
      {
        "type": "NOT",
        "assertion": {
          "type": "EXPECT_DISPLAY_VALUE",
          "displayValue": "0"
        }
      }
    ]
  }
}
```

For a runnable mutate-then-verify plan, see
[`examples/assertion-request.json`](../examples/assertion-request.json).

---

## Inspection Queries

Inspection queries are ordered, explicit post-mutation or post-assertion requests. Every
inspection must include a caller-defined `stepId`, and every result echoes that `stepId` back in
the successful response.

Inspection categories:

- Introspection: exact workbook facts with no higher-level interpretation.
- Analysis: finding-bearing workbook conclusions built on top of introspection.

```json
{
  "steps": [
    {
      "stepId": "workbook",
      "target": {
        "type": "CURRENT"
      },
      "query": {
        "type": "GET_WORKBOOK_SUMMARY"
      }
    },
    {
      "stepId": "inventory-window",
      "target": {
        "type": "RECTANGULAR_WINDOW",
        "sheetName": "Inventory",
        "topLeftAddress": "A1",
        "rowCount": 5,
        "columnCount": 3
      },
      "query": {
        "type": "GET_WINDOW"
      }
    },
    {
      "stepId": "inventory-schema",
      "target": {
        "type": "RECTANGULAR_WINDOW",
        "sheetName": "Inventory",
        "topLeftAddress": "A1",
        "rowCount": 5,
        "columnCount": 3
      },
      "query": {
        "type": "GET_SHEET_SCHEMA"
      }
    }
  ]
}
```

### GET_WORKBOOK_SUMMARY

Returns workbook-level summary facts such as sheet order, named-range count, and the workbook
force-recalculation flag.

```json
{
  "stepId": "workbook",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "GET_WORKBOOK_SUMMARY"
  }
}
```

Response shapes:

- `{"kind":"EMPTY","sheetCount":0,"sheetNames":[],"namedRangeCount":0,"forceFormulaRecalculationOnOpen":false}`
- `{"kind":"WITH_SHEETS","sheetCount":2,"sheetNames":["Budget","Budget Review"],"activeSheetName":"Budget Review","selectedSheetNames":["Budget","Budget Review"],"namedRangeCount":0,"forceFormulaRecalculationOnOpen":false}`

`selectedSheetNames` are returned in workbook order, not request order.

### GET_PACKAGE_SECURITY

Returns factual OOXML package-security state for the currently open workbook: whether the package
is encrypted, which package-encryption mode is present, and the validation state of each OOXML
package signature.

```json
{
  "stepId": "security",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "GET_PACKAGE_SECURITY"
  }
}
```

Response shape:

```json
{
  "security": {
    "encryption": {
      "encrypted": true,
      "mode": "AGILE",
      "cipherAlgorithm": "aes",
      "hashAlgorithm": "sha512",
      "chainingMode": "ChainingModeCBC",
      "keyBits": 256,
      "blockSize": 16,
      "spinCount": 100000
    },
    "signatures": [
      {
        "packagePartName": "/_xmlsignatures/sig1.xml",
        "signerSubject": "CN=GridGrind Signing",
        "signerIssuer": "CN=GridGrind Signing",
        "serialNumberHex": "01",
        "state": "VALID"
      }
    ]
  }
}
```

`GET_PACKAGE_SECURITY` runs only on the full-XSSF read path. `execution.mode.readMode=EVENT_READ`
rejects it up front because the event-model reader exposes only workbook and sheet summaries.
Unencrypted workbooks return `"encryption": { "encrypted": false }` plus an empty `signatures`
array.

### GET_WORKBOOK_PROTECTION

Returns workbook-level protection facts such as structure, windows, and revisions lock state plus
whether the workbook stores password hashes for the workbook or revisions protection domains.

```json
{
  "stepId": "workbook-protection",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "GET_WORKBOOK_PROTECTION"
  }
}
```

Response shape:

```json
{
  "protection": {
    "structureLocked": false,
    "windowsLocked": false,
    "revisionsLocked": false,
    "workbookPasswordHashPresent": false,
    "revisionsPasswordHashPresent": false
  }
}
```

### GET_CUSTOM_XML_MAPPINGS

Return workbook custom-XML mapping metadata, including identifiers, schema metadata, linked
single cells, linked tables, and optional data-binding facts. The step target is always
`WorkbookSelector.CURRENT`.

```json
{
  "stepId": "custom-xml-mappings",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "GET_CUSTOM_XML_MAPPINGS"
  }
}
```

Response shape:

```json
{
  "mappings": [
    {
      "mapId": 1,
      "name": "CORSO_mapping",
      "rootElement": "CORSO",
      "schemaId": "Schema1",
      "linkedCells": [
        {
          "sheetName": "Foglio1",
          "address": "A1",
          "xpath": "/CORSO/NOME",
          "xmlDataType": "string"
        }
      ]
    }
  ]
}
```

### EXPORT_CUSTOM_XML_MAPPING

Export one existing workbook custom-XML mapping as serialized XML. The step target is always
`WorkbookSelector.CURRENT`.

```json
{
  "stepId": "custom-xml-export",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "EXPORT_CUSTOM_XML_MAPPING",
    "mapping": {
      "mapId": 1,
      "name": "CORSO_mapping"
    },
    "validateSchema": true,
    "encoding": "UTF-8"
  }
}
```

Response shape:

```json
{
  "export": {
    "encoding": "UTF-8",
    "schemaValidated": true,
    "xml": "<CORSO><NOME>Grid</NOME></CORSO>"
  }
}
```

### GET_NAMED_RANGES

Returns exact named-range reports selected by workbook-wide or exact-selector input.

```json
{
  "stepId": "named-ranges",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "GET_NAMED_RANGES"
  }
}
```

```json
{
  "stepId": "selected-named-ranges",
  "target": {
    "type": "ANY_OF",
    "selectors": [
      {
        "type": "WORKBOOK_SCOPE",
        "name": "BudgetTotal"
      },
      {
        "type": "SHEET_SCOPE",
        "sheetName": "Budget",
        "name": "BudgetTable"
      }
    ]
  },
  "query": {
    "type": "GET_NAMED_RANGES"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `stepId` | Yes | Stable caller-defined correlation identifier. |
| `target` | Yes | `ALL`, `BY_NAME`, `BY_NAMES`, `WORKBOOK_SCOPE`, `SHEET_SCOPE`, or `ANY_OF` named-range selector. |

Selected named-range reads fail with `NAMED_RANGE_NOT_FOUND` when any selector does not match an
existing named range.

### GET_SHEET_SUMMARY

Returns structural summary facts for one sheet.

```json
{
  "stepId": "sheet-summary",
  "target": {
    "type": "BY_NAME",
    "name": "Inventory"
  },
  "query": {
    "type": "GET_SHEET_SUMMARY"
  }
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
  "stepId": "cells",
  "target": {
    "type": "BY_ADDRESSES",
    "sheetName": "Inventory",
    "addresses": [
      "A1",
      "B4",
      "C10"
    ]
  },
  "query": {
    "type": "GET_CELLS"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `stepId` | Yes | Stable caller-defined correlation identifier. |
| `target` | Yes | `BY_ADDRESSES` cell selector carrying the owning sheet plus ordered A1 addresses. |

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

### GET_ARRAY_FORMULAS

Return factual array-formula group metadata for the selected sheets.

```json
{
  "stepId": "array-formulas",
  "target": {
    "type": "BY_NAME",
    "name": "Calc"
  },
  "query": {
    "type": "GET_ARRAY_FORMULAS"
  }
}
```

Response shape:

```json
{
  "arrayFormulas": [
    {
      "sheetName": "Calc",
      "range": "D2:D4",
      "topLeftAddress": "D2",
      "formula": "B2:B4*C2:C4",
      "singleCell": false
    }
  ]
}
```

Each entry reports the stored contiguous array-formula range, the top-left anchor cell, the
normalized formula text without a leading `=`, and whether the stored group is single-cell.

### GET_WINDOW

Returns a rectangular top-left-anchored window of cell snapshots. The window includes styled blank
cells so template-like workbooks remain visible.

`rowCount * columnCount` must not exceed 250,000. The window must not extend beyond the Excel
2007 sheet boundary (rows 0–1,048,575, columns 0–16,383); requests that overflow are rejected
with `INVALID_REQUEST`.

```json
{
  "stepId": "window",
  "target": {
    "type": "RECTANGULAR_WINDOW",
    "sheetName": "Inventory",
    "topLeftAddress": "A1",
    "rowCount": 5,
    "columnCount": 3
  },
  "query": {
    "type": "GET_WINDOW"
  }
}
```

Response shape: `{ "window": { "sheetName": "...", "rows": [ { "cells": [...] } ] } }`. The
top-level key is `window` and cells are nested under `window.rows[N].cells`. This differs from
`GET_CELLS` where cells are directly under the top-level `cells` key.

### GET_MERGED_REGIONS

Returns the exact merged regions defined on one sheet.

```json
{
  "stepId": "merged",
  "target": {
    "type": "BY_NAME",
    "name": "Inventory"
  },
  "query": {
    "type": "GET_MERGED_REGIONS"
  }
}
```

### GET_HYPERLINKS

Returns hyperlink metadata for selected cells on one sheet. Response hyperlinks reuse the same
discriminated shape as `SET_HYPERLINK` targets. `FILE` targets come back in the `path` field as
normalized plain path strings, not `file:` URIs.

```json
{
  "stepId": "hyperlinks",
  "target": {
    "type": "ALL_USED_IN_SHEET",
    "sheetName": "Inventory"
  },
  "query": {
    "type": "GET_HYPERLINKS"
  }
}
```

```json
{
  "stepId": "selected-hyperlinks",
  "target": {
    "type": "BY_ADDRESSES",
    "sheetName": "Inventory",
    "addresses": [
      "A1",
      "B4"
    ]
  },
  "query": {
    "type": "GET_HYPERLINKS"
  }
}
```

### GET_COMMENTS

Returns comment metadata for selected cells on one sheet.

```json
{
  "stepId": "comments",
  "target": {
    "type": "ALL_USED_IN_SHEET",
    "sheetName": "Inventory"
  },
  "query": {
    "type": "GET_COMMENTS"
  }
}
```

```json
{
  "stepId": "selected-comments",
  "target": {
    "type": "BY_ADDRESSES",
    "sheetName": "Inventory",
    "addresses": [
      "A1",
      "B4"
    ]
  },
  "query": {
    "type": "GET_COMMENTS"
  }
}
```

Cell-selector payloads use:

```json
{
  "type": "ALL_USED_IN_SHEET",
  "sheetName": "Inventory"
}
{
  "type": "BY_ADDRESSES",
  "sheetName": "Inventory",
  "addresses": [
    "A1",
    "B4"
  ]
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
  "stepId": "drawing-objects",
  "target": {
    "type": "ALL_ON_SHEET",
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
    "type": "ALL_ON_SHEET",
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
    "type": "BY_NAME",
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
    "type": "BY_NAME",
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
    "type": "BY_NAME",
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
- `tabColor`: `null` or a structured color report
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
    "type": "BY_NAME",
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
    "type": "ALL_ON_SHEET",
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
    "type": "BY_RANGES",
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
  "type": "ALL_ON_SHEET",
  "sheetName": "Inventory"
}
{
  "type": "BY_RANGES",
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
    "type": "ALL_ON_SHEET",
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
    "type": "BY_RANGES",
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
    "type": "BY_NAME",
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
    "type": "ALL"
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
    "type": "BY_NAMES",
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
  "type": "ALL"
}
{
  "type": "BY_NAMES",
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
    "type": "ALL"
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
    "type": "BY_NAMES",
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
    "type": "ALL"
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
    "type": "BY_NAMES",
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
  "type": "ALL"
}
{
  "type": "BY_NAMES",
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

`dominantType` per column is `null` when all data cells are blank, or when two or more types tie
for the highest count. Formula cells contribute their evaluated result type (e.g. `NUMBER`,
`STRING`) to `observedTypes` and `dominantType`, not `FORMULA`.

```json
{
  "stepId": "schema",
  "target": {
    "type": "RECTANGULAR_WINDOW",
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
    "type": "ALL"
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
    "type": "ANY_OF",
    "selectors": [
      {
        "type": "WORKBOOK_SCOPE",
        "name": "BudgetTotal"
      },
      {
        "type": "SHEET_SCOPE",
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

### ANALYZE_FORMULA_HEALTH

Reports finding-bearing formula health across one or more sheets. This is where volatile
functions, formula-error results, and evaluation failures surface.

```json
{
  "stepId": "formula-health",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "ANALYZE_FORMULA_HEALTH"
  }
}
```

Response shape: `{ "analysis": { "checkedFormulaCellCount": ..., "summary": ..., "findings":`
`[...] } }`.

### ANALYZE_DATA_VALIDATION_HEALTH

Reports finding-bearing data-validation health across one or more sheets. Findings include
unsupported rules, broken formulas, and overlapping validation coverage.

```json
{
  "stepId": "data-validation-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "ANALYZE_DATA_VALIDATION_HEALTH"
  }
}
```

### ANALYZE_CONDITIONAL_FORMATTING_HEALTH

Reports conditional-formatting findings such as broken formulas, unsupported loaded rule
families, empty target ranges, or priority collisions.

```json
{
  "stepId": "conditional-formatting-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"
  }
}
```

### ANALYZE_AUTOFILTER_HEALTH

Reports autofilter findings such as invalid ranges, blank header rows, or ownership mismatches
between sheet-level filters and tables.

```json
{
  "stepId": "autofilter-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "ANALYZE_AUTOFILTER_HEALTH"
  }
}
```

### ANALYZE_TABLE_HEALTH

Reports table findings such as overlaps, broken ranges, blank or duplicate headers, or style
mismatches.

```json
{
  "stepId": "table-health",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "ANALYZE_TABLE_HEALTH"
  }
}
```

```json
{
  "stepId": "selected-table-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "InventoryTable",
      "Trips"
    ]
  },
  "query": {
    "type": "ANALYZE_TABLE_HEALTH"
  }
}
```

### ANALYZE_PIVOT_TABLE_HEALTH

Reports pivot-table findings such as missing cache parts, missing workbook-cache definitions,
broken sources, duplicate names, synthetic fallback names, or unsupported persisted detail.

```json
{
  "stepId": "pivot-health",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "ANALYZE_PIVOT_TABLE_HEALTH"
  }
}
```

```json
{
  "stepId": "selected-pivot-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "SalesPivot",
      "NamedPivot"
    ]
  },
  "query": {
    "type": "ANALYZE_PIVOT_TABLE_HEALTH"
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
  "stepId": "hyperlink-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "ANALYZE_HYPERLINK_HEALTH"
  }
}
```

### ANALYZE_NAMED_RANGE_HEALTH

Reports named-range findings such as broken references, unresolved targets, and scope shadowing.

```json
{
  "stepId": "named-range-health",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "ANALYZE_NAMED_RANGE_HEALTH"
  }
}
```

Response shape: `{ "analysis": { "checkedNamedRangeCount": ..., "summary": ..., "findings":`
`[...] } }`.

### ANALYZE_WORKBOOK_FINDINGS

Runs every shipped finding-bearing analysis family across the workbook and returns one aggregated
flat finding list. This is the primary workbook-health check and works especially well with
`persistence.type=NONE` when you want a no-save lint pass.

Response shape: `{ "analysis": { "summary": ..., "findings": [...] } }`.

The aggregate currently includes formula, data validation, conditional formatting, autofilter,
table, pivot-table, hyperlink, and named-range findings.

```json
{
  "stepId": "workbook-findings",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "ANALYZE_WORKBOOK_FINDINGS"
  }
}
```

Formula references to same-request sheet names with spaces should use single quotes, for example
`'Budget Review'!B1`. When execution succeeds, GridGrind reports this style of request issue in
the success response `warnings` array.
