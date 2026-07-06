---
afad: "4.0"
version: "0.72.0"
domain: WORKBOOK_CELL_INSPECTIONS
updated: "2026-07-02"
route:
  keywords: [gridgrind, inspections, get-workbook-summary, get-package-security, get-cells, get-window, get-comments]
  questions: ["how do i inspect workbook facts in gridgrind", "how do i read cells in gridgrind", "how do i inspect package security in gridgrind"]
---

# Workbook And Cell Inspection Reference

**Purpose**: Detailed reference for workbook-core, sheet-core, cell, window, hyperlink, and
comment inspection queries.
**Landing page**: [ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[ASSERTIONS.md](./ASSERTIONS.md), [DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md),
and [ANALYSIS_QUERIES.md](./ANALYSIS_QUERIES.md)

### GET_WORKBOOK_SUMMARY

Returns workbook-level summary facts such as sheet order, named-range count, and the workbook
force-recalculation flag.

```json
{
  "stepId": "workbook",
  "target": {
    "type": "WORKBOOK_CURRENT"
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
    "type": "WORKBOOK_CURRENT"
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
      "type": "ENCRYPTED",
      "mode": "AGILE",
      "cipherAlgorithm": "AES_256",
      "hashAlgorithm": "SHA_512",
      "chainingMode": "CBC",
      "keyBits": 256,
      "blockSize": 16,
      "spinCount": 100000
    },
    "signatures": [
      {
        "packagePartName": "/_xmlsignatures/sig1.xml",
        "signer": {
          "subject": "CN=GridGrind Signing",
          "issuer": "CN=GridGrind Signing",
          "serialNumberHex": "01"
        },
        "state": "VALID"
      }
    ]
  }
}
```

`GET_PACKAGE_SECURITY` runs only on the full-XSSF read path. `execution.mode.type=EVENT_READ`
rejects it up front because the event-model reader exposes only workbook and sheet summaries.
Unencrypted workbooks return `"encryption": { "type": "NONE" }` plus an empty `signatures`
array. Readback continues to report factual legacy modes such as `STANDARD` when the loaded OOXML
package uses them, even though the write-side request contract authors AGILE packages only.

### GET_WORKBOOK_PROTECTION

Returns workbook-level protection facts such as structure, windows, and revisions lock state plus
whether the workbook stores password hashes for the workbook or revisions protection domains.

```json
{
  "stepId": "workbook-protection",
  "target": {
    "type": "WORKBOOK_CURRENT"
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
    "type": "WORKBOOK_CURRENT"
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
    "type": "WORKBOOK_CURRENT"
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
    "type": "NAMED_RANGE_ALL"
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
    "type": "SHEET_BY_NAME",
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
    "type": "CELL_BY_ADDRESSES",
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

`addresses` must not be empty, must not contain duplicates, and must not contain more than
250,000 entries. This uses the same deterministic read-size cap as `GET_WINDOW` and
`GET_SHEET_SCHEMA`, but for `GET_CELLS` the cap keys off the exact returned cell count rather than
rectangular area.

`GET_CELLS`, `GET_WINDOW`, and `GET_SHEET_SCHEMA` all accept the same optional
`query.projection.facets` selector. Omit `projection` for the default compact `[VALUE]` readback,
or add only the extra facets you need:

```json
{
  "type": "GET_CELLS",
  "projection": {
    "facets": [
      "VALUE",
      "FORMAT",
      "TEMPORAL"
    ]
  }
}
```

Under that default `[VALUE]` projection, date-like numeric cells still read back as
`type=NUMBER` plus `numberValue`. Add `TEMPORAL` whenever you need GridGrind to distinguish a
plain number from a date, time, or date-time value.

#### Cell snapshot shape

Every factual cell in a `GET_CELLS` or `GET_WINDOW` response always includes these canonical
fields:

| Field | Description |
|:------|:------------|
| `address` | A1-style cell address. |
| `type` | Published readback discriminator: `BLANK`, `TEXT`, `NUMBER`, `BOOLEAN`, `ERROR`, or `FORMULA`. |

There is no separate readback `RICH_TEXT` discriminator. Authored rich text reads back as
`type=TEXT` with optional `runs` when the `RICH_TEXT_RUNS` facet is projected.

Facet-gated fields:

| Facet | Fields surfaced |
|:------|:----------------|
| `VALUE` | `textValue`, `numberValue`, `booleanValue`, `errorValue`, and formula `evaluation` |
| `FORMAT` | `displayValue` |
| `STYLE` | `style` |
| `HYPERLINK` | `hyperlink` |
| `COMMENT` | `comment` |
| `FORMULA` | `formula` |
| `RICH_TEXT_RUNS` | text-cell or evaluated-text `runs` |
| `TEMPORAL` | numeric `temporal` |

The protocol catalog publishes this same mapping in machine-readable form: field descriptors on
`nestedTypes:cellReportTypes` and `nestedTypes:cellValueReportTypes` carry `projectedByFacets`,
so an agent can derive `FORMAT -> displayValue`, `FORMULA -> formula`,
`RICH_TEXT_RUNS -> runs`, and `TEMPORAL -> temporal` directly from the catalog payload. The
compact protocol-catalog index also publishes a legend for field-metadata keys such as
`projectedByFacets`, `noteRefs`, and `enumValueDocs`, while `plainTypes:cellReadProjectionType`
uses `enumValueDocs` on `facets` so tokens such as `VALUE`, `FORMAT`, and `TEMPORAL` explain
their own output directly in the machine contract. Shared catalog rules such as request-owned
path resolution now appear once under top-level `notes`, with entry-local `noteRefs` pointing at
the stable note id instead of repeating the same paragraph across summaries.

Type-specific fields:

| Field | Present when | Description |
|:------|:-------------|:------------|
| `textValue` | `type=TEXT` and `VALUE` projected | The stored plain string content. |
| `runs` | `type=TEXT` and `RICH_TEXT_RUNS` projected | Optional ordered rich-text runs. When present, run text concatenates exactly to `textValue`. |
| `numberValue` | `type=NUMBER` and `VALUE` projected | The numeric value as a double. |
| `temporal` | `type=NUMBER` and `TEMPORAL` projected | Optional temporal facet `{ isDate, kind, isoValue }`. Date-like numeric cells stay `type=NUMBER`; `kind` refines them to `DATE`, `TIME`, or `DATE_TIME`. |
| `booleanValue` | `type=BOOLEAN` and `VALUE` projected | `true` or `false`. |
| `errorValue` | `type=ERROR` and `VALUE` projected | The GridGrind-owned reported error literal. Stored error cells use canonical Excel tokens such as `#DIV/0!` and `#REF!`; evaluated states may additionally report `#CIRCULAR_REF!` or `#FUNCTION_NOT_IMPLEMENTED!`. |
| `formula` | `type=FORMULA` and `FORMULA` projected | The stored formula text without the leading `=`. |
| `evaluation` | `type=FORMULA` and `VALUE` projected | Nested `CellValueReport` carrying the evaluated value kind (`BLANK`, `TEXT`, `NUMBER`, `BOOLEAN`, or `ERROR`). |
| `displayValue` | `FORMAT` projected | Formatted string exactly as Excel would render it. |
| `style` | `STYLE` projected | Nested style snapshot with `numberFormat`, `alignment`, `font`, `fill`, `border`, and `protection`. |
| `hyperlink` | `HYPERLINK` projected and metadata exists | Optional hyperlink metadata attached at snapshot time. |
| `comment` | `COMMENT` projected and metadata exists | Optional comment metadata including plain text, visibility, optional runs, and optional anchor bounds. |

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
style contract now uses the same structured color families via `ColorInput` plus explicit fill
variants under `style.fill.kind`. Agents that read a `CellColorReport` and want to write it back
can preserve the same color semantics, but they must still translate read-side report objects into
the matching write-side request types and choose the correct `style.fill.kind` variant instead of
echoing the entire read-back cell style verbatim.

### GET_ARRAY_FORMULAS

Return factual array-formula group metadata for the selected sheets.

```json
{
  "stepId": "array-formulas",
  "target": {
    "type": "SHEET_BY_NAME",
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

Returns a rectangular top-left-anchored window of cell snapshots. `GET_WINDOW` is sparse by
default: omitted blank cells do not appear in the payload at all, even when Excel still carries
style on them. Omit `includeBlanks` for that sparse default, or set `query.includeBlanks=true`
when you need the explicit dense row grid.

`rowCount * columnCount` must not exceed 250,000. All three cell-returning read surfaces share
this deterministic limit: `GET_CELLS` counts explicit addresses, while `GET_WINDOW` and
`GET_SHEET_SCHEMA` count requested rectangular area. Sparse windows collapse blank cells in the
response, but dense data can still fill the whole rectangle and GridGrind rejects over-large
windows before workbook IO. The window must not extend beyond the Excel 2007 sheet boundary (rows
0–1,048,575, columns 0–16,383); requests that overflow are rejected with `INVALID_REQUEST`.

```json
{
  "stepId": "window",
  "target": {
    "type": "RANGE_RECTANGULAR_WINDOW",
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

Response shapes:

- Sparse default:
  `{ "window": { "shape": "SPARSE", "sheetName": "...", "topLeftAddress": "A1", "dimensions": { "rowCount": 5, "columnCount": 3 }, "populatedCells": [...] } }`
- Dense with `includeBlanks=true`:
  `{ "window": { "shape": "DENSE", "sheetName": "...", "topLeftAddress": "A1", "dimensions": { "rowCount": 5, "columnCount": 3 }, "rows": [ { "rowIndex": 0, "cells": [...] } ] } }`

`GET_CELLS` returns its cell list directly under top-level `cells`; `GET_WINDOW` always nests the
window payload under top-level `window`.

### GET_MERGED_REGIONS

Returns the exact merged regions defined on one sheet.

```json
{
  "stepId": "merged",
  "target": {
    "type": "SHEET_BY_NAME",
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
    "type": "CELL_ALL_USED_IN_SHEET",
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
    "type": "CELL_BY_ADDRESSES",
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
    "type": "CELL_ALL_USED_IN_SHEET",
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
    "type": "CELL_BY_ADDRESSES",
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
  "type": "CELL_ALL_USED_IN_SHEET",
  "sheetName": "Inventory"
}
{
  "type": "CELL_BY_ADDRESSES",
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
