---
afad: "3.5"
version: "0.50.0"
domain: OPERATIONS
updated: "2026-04-21"
route:
  keywords: [gridgrind, operations, assertions, reads, introspection, analysis, assert, expect-cell-value, expect-analysis-max-severity, set-cell, set-range, apply-style, ensure-sheet, rename-sheet, delete-sheet, move-sheet, copy-sheet, set-active-sheet, set-selected-sheets, set-sheet-visibility, set-sheet-protection, clear-sheet-protection, set-workbook-protection, clear-workbook-protection, merge-cells, unmerge-cells, set-column-width, set-row-height, set-sheet-pane, set-sheet-zoom, set-print-layout, clear-print-layout, freeze-panes, split-panes, set-data-validation, set-autofilter, clear-autofilter, set-table, delete-table, set-pivot-table, delete-pivot-table, set-picture, set-chart, set-shape, set-embedded-object, set-signature-line, set-drawing-object-anchor, delete-drawing-object, append-row, clear-range, execution-calculation, evaluate-all, evaluate-targets, clear-caches-only, markrecalculateonopen, auto-size-columns, get-cells, get-window, get-print-layout, get-package-security, get-workbook-protection, get-data-validations, get-autofilters, get-tables, get-pivot-tables, get-drawing-objects, get-charts, get-drawing-object-payload, get-sheet-layout, get-sheet-schema, analyze-autofilter-health, analyze-table-health, analyze-pivot-table-health, analyze-workbook-findings, source-backed, inline, utf8_file, standard_input, inline_base64, file, ooxml, package-security, encryption, signing, request, json, protocol, coordinates, rowindex, columnindex, warnings]
  questions: ["what operations does gridgrind support", "what assertions does gridgrind support", "what reads does gridgrind support", "how do I assert a cell value in gridgrind", "how do I assert workbook health in gridgrind", "how do I rename a sheet", "how do I delete a sheet", "how do I move a sheet", "how do I copy a sheet", "how do I set the active sheet", "how do I set selected sheets", "how do I set sheet visibility", "how do I set sheet protection", "how do I set workbook protection", "how do I inspect package security in gridgrind", "how do I open an encrypted workbook in gridgrind", "how do I sign a workbook in gridgrind", "how do I merge cells", "how do I set a column width", "how do I freeze panes", "how do I set split panes", "how do I set sheet zoom", "how do I set print layout", "how do I set a cell value", "how do I apply a style", "how do I write a range", "how do I create an autofilter in gridgrind", "how do I create a table in gridgrind", "how do I create a pivot table in gridgrind", "how do I read pivot tables in gridgrind", "how do I add a picture in gridgrind", "how do I add a signature line in gridgrind", "how do I author a chart in gridgrind", "how do I read charts in gridgrind", "how do I read drawing objects in gridgrind", "what is the request format", "what fields does SET_RANGE accept", "what does GET_CELLS accept", "which fields use A1 notation versus zero-based indexes", "how do I run workbook findings without saving", "how do I load long text from a file in gridgrind", "how do I send binary payloads to gridgrind"]
---

# Operations, Assertions, and Inspection Reference

**Purpose**: Complete reference for all GridGrind request fields, mutation actions, assertion
types, and inspection queries.
**Prerequisites**: Familiarity with the [README](../README.md) quick start.
**Limits**: See [LIMITATIONS.md](LIMITATIONS.md) for all hard ceilings — window cell count, Excel maximums, and memory constraints.

The CLI can emit the current minimal request and the full machine-readable protocol inventory:

```bash
gridgrind --print-request-template
gridgrind --print-protocol-catalog
gridgrind --print-example ASSERTION
```

The protocol catalog is designed for black-box consumers. Each published type entry includes field
descriptors with required/optional status and recursive shape metadata, so fields like `target`,
`action`, `query`, `value`, `style`, and `scope` point directly at the nested/plain type group
that defines their accepted JSON structure. Mutation, assertion, and inspection entries also
publish `targetSelectors` or `targetSelectorRule` so black-box callers can see the exact allowed
target families before sending a request. `--print-example <id>` emits one built-in generated
request from the same contract-owned registry that backs the committed `examples/*.json` files.

---

## Source-Backed Authored Inputs

GridGrind's mutation contract is source-backed for large text and binary payloads. The request
body stays canonical JSON, while the authored value can come from inline literals, files in the
execution environment, or bound stdin bytes.

Text-bearing mutation fields use `TextSourceInput`:

```json
{ "type": "INLINE", "text": "Quarterly note" }
{ "type": "UTF8_FILE", "path": "authored-inputs/note.txt" }
{ "type": "STANDARD_INPUT" }
```

Binary-bearing mutation fields use `BinarySourceInput`:

```json
{ "type": "INLINE_BASE64", "base64Data": "SGVsbG8=" }
{ "type": "FILE", "path": "authored-inputs/payload.bin" }
{ "type": "STANDARD_INPUT" }
```

- Relative `UTF8_FILE` and `FILE` paths resolve from the current working directory.
- `STANDARD_INPUT` authored values require `--request <path>` on the CLI because stdin cannot
  carry both the request JSON and authored input content in the same invocation.
- The request JSON transport is capped at 16 MiB. Large authored text and binary payloads belong
  in `UTF8_FILE`, `FILE`, or `STANDARD_INPUT` sources instead of inline JSON strings.
- Source-backed input loading happens before workbook open and is journaled under
  `journal.inputResolution`.
- Source-backed failures classify as `INPUT_SOURCE_NOT_FOUND`, `INPUT_SOURCE_UNAVAILABLE`, or
  `INPUT_SOURCE_IO_ERROR`.
- See [`examples/source-backed-input-request.json`](../examples/source-backed-input-request.json)
  for a complete file-backed text, formula, and binary request.

---

## Request Structure

```json
{
  "protocolVersion": "V1",
  "source":      { ... },
  "persistence": { ... },
  "execution": { ... },
  "formulaEnvironment": { ... },
  "steps": [ ... ]
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `protocolVersion` | No | Wire-contract version. Defaults to `V1`. Include it — future breaking revisions will be explicit. |
| `source` | Yes | Where the workbook comes from. |
| `persistence` | No | Where and whether to save. Omit to run steps without saving. |
| `execution` | No | Optional execution policy for low-memory mode selection, structured journaling, and formula calculation handling. Omit for the default full-XSSF path with `NORMAL` journaling and `DO_NOT_CALCULATE`. |
| `formulaEnvironment` | No | Request-scoped evaluator configuration for external workbook bindings, missing-workbook policy, and template-backed UDF toolpacks. |
| `steps` | No | Ordered list of workbook mutations, assertions, and inspections. |

Every tagged request union uses `type` as its discriminator field: `source`, `persistence`,
`action`, `query`, cell values, hyperlink targets, selectors, and named-range scopes.
Every step object carries exactly one of `action`, `assertion`, or `query`. Step kind is inferred
from that field; request steps do not carry a separate `step.type`.

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

### Execution Policy

`execution` is optional. Omit it for the default `FULL_XSSF` request path with `NORMAL`
journaling.

```json
{
  "execution": {
    "mode": {
      "readMode": "EVENT_READ",
      "writeMode": "STREAMING_WRITE"
    },
    "journal": {
      "level": "VERBOSE"
    },
    "calculation": {
      "strategy": {
        "type": "DO_NOT_CALCULATE"
      },
      "markRecalculateOnOpen": true
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `mode` | No | Optional low-memory read and write mode selection. Omit it for the default full-XSSF path. |
| `journal` | No | Optional structured-journal policy. Omit it for `NORMAL` detail. |
| `calculation` | No | Optional formula-calculation policy covering immediate evaluation, cache clearing, and workbook-open recalc flags. |

- `execution.mode.readMode: EVENT_READ` selects the low-memory XSSF event-model reader. It supports only
  `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY` (`LIM-019`).
- `execution.mode.writeMode: STREAMING_WRITE` selects the low-memory SXSSF writer. It requires `source.type:
  NEW`, supports only `ENSURE_SHEET` and `APPEND_ROW`,
  requires `execution.calculation.strategy=DO_NOT_CALCULATE`,
  allows `markRecalculateOnOpen=true`, and
  requires at least one `ENSURE_SHEET` or `APPEND_ROW` (`LIM-020`).
- `execution.journal.level` accepts `SUMMARY`, `NORMAL`, and `VERBOSE`.
- `execution.calculation.strategy` accepts `DO_NOT_CALCULATE`, `EVALUATE_ALL`,
  `EVALUATE_TARGETS`, and `CLEAR_CACHES_ONLY`.
- `execution.calculation.markRecalculateOnOpen` persists Excel's workbook-level recalc-on-open
  flag without requiring an extra mutation step.
- `VERBOSE` keeps the structured response journal and also streams fine-grained execution events to
  CLI stderr while the request is running.
- `EVENT_READ` can run directly against an existing workbook when the request is read-only and
  unsaved. If the request also performs full-XSSF mutations, GridGrind materializes the mutated
  workbook state and then performs the summary reads through the event model.
- `STREAMING_WRITE` can pair with either `readMode: FULL_XSSF` for broader readback or
  `readMode: EVENT_READ` for summary-only low-memory readback from the materialized streaming
  result.

### Response Journal

Every success and failure response includes a structured `journal` object:

```json
{
  "journal": {
    "planId": "budget-pass",
    "level": "NORMAL",
    "source": {
      "type": "EXISTING",
      "path": "budget.xlsx"
    },
    "persistence": {
      "type": "SAVE_AS",
      "path": "out/budget-reviewed.xlsx"
    },
    "validation": {
      "status": "SUCCEEDED",
      "startedAt": "2026-04-19T09:30:00Z",
      "finishedAt": "2026-04-19T09:30:00Z",
      "durationMillis": 1
    },
    "inputResolution": {
      "status": "SUCCEEDED",
      "startedAt": "2026-04-19T09:30:00Z",
      "finishedAt": "2026-04-19T09:30:00Z",
      "durationMillis": 0
    },
    "open": {
      "status": "SUCCEEDED",
      "startedAt": "2026-04-19T09:30:00Z",
      "finishedAt": "2026-04-19T09:30:09Z",
      "durationMillis": 9
    },
    "calculation": {
      "preflight": {
        "status": "SUCCEEDED",
        "startedAt": "2026-04-19T09:30:09Z",
        "finishedAt": "2026-04-19T09:30:10Z",
        "durationMillis": 1
      },
      "execution": {
        "status": "SUCCEEDED",
        "startedAt": "2026-04-19T09:30:10Z",
        "finishedAt": "2026-04-19T09:30:12Z",
        "durationMillis": 2
      }
    },
    "persistencePhase": {
      "status": "SUCCEEDED",
      "startedAt": "2026-04-19T09:30:14Z",
      "finishedAt": "2026-04-19T09:30:28Z",
      "durationMillis": 14
    },
    "close": {
      "status": "SUCCEEDED",
      "startedAt": "2026-04-19T09:30:28Z",
      "finishedAt": "2026-04-19T09:30:29Z",
      "durationMillis": 1
    },
    "steps": [
      {
        "stepIndex": 0,
        "stepId": "set-total",
        "stepKind": "MUTATION",
        "stepType": "SET_CELL",
        "resolvedTargets": [
          {
            "kind": "CELL",
            "label": "Cell Budget!B2"
          }
        ],
        "phase": {
          "status": "SUCCEEDED",
          "startedAt": "2026-04-19T09:30:12Z",
          "finishedAt": "2026-04-19T09:30:14Z",
          "durationMillis": 2
        },
        "outcome": "SUCCEEDED"
      }
    ],
    "warnings": [],
    "outcome": {
      "status": "SUCCEEDED",
      "plannedStepCount": 1,
      "completedStepCount": 1,
      "durationMillis": 29
    }
  }
}
```

| Field | Description |
|:------|:------------|
| `planId` | Caller-supplied plan correlation ID when present, otherwise a synthesized internal ID after request parsing. Pre-parse CLI failures may omit it because no request plan was available yet. |
| `level` | `SUMMARY`, `NORMAL`, or `VERBOSE`, matching the effective `execution.journal.level`. |
| `validation`, `inputResolution`, `open`, `persistencePhase`, `close` | Top-level pipeline phase summaries. Finished phases carry `status`, `startedAt`, `finishedAt`, and `durationMillis`; `NOT_STARTED` phases carry `status` plus `durationMillis=0` only. `inputResolution` records source-backed file/stdin loading before workbook open. |
| `calculation` | Top-level calculation telemetry. `preflight` classifies authored formulas and `execution` records the requested evaluation or cache-clearing work. |
| `steps[]` | Ordered per-step telemetry including `resolvedTargets`, phase timing, outcome, and optional failure classification. `resolvedTargets` is compact in `SUMMARY` and expanded in `NORMAL`/`VERBOSE`. |
| `warnings[]` | Request-phase warnings derived during execution, mirrored from the top-level success `warnings` array. |
| `events[]` | Fine-grained live events. Present only when `level=VERBOSE`. |
| `outcome` | Whole-run status plus `plannedStepCount`, `completedStepCount`, total `durationMillis`, and optional `failedStepIndex`, `failedStepId`, and `failureCode` when the run failed. |

`VERBOSE` keeps the full response journal and also streams `events[]` entries live to CLI stderr
while the request is running.

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
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "NONE"
  },
  "steps": [
    {
      "stepId": "lint",
      "target": {
        "type": "CURRENT"
      },
      "query": {
        "type": "ANALYZE_WORKBOOK_FINDINGS"
      }
    }
  ]
}
```

Successful responses may include a `warnings` array. The current request-phase warning flags
same-request sheet names with spaces referenced in formulas without single quotes. Use
`'Sheet Name'!A1` syntax for those references.

For batch health-plus-read workflows, see
[`examples/workbook-health-request.json`](../examples/workbook-health-request.json) for a compact
no-save pass and
[`examples/introspection-analysis-request.json`](../examples/introspection-analysis-request.json)
for a broader mixed introspection and analysis run.

---

## Source

```json
{
  "type": "NEW"
}
```
Create a new blank `.xlsx` workbook. The workbook starts with zero sheets; use `ENSURE_SHEET` to
create the first sheet before writing any cells.

```json
{
  "type": "EXISTING",
  "path": "path/to/workbook.xlsx"
}
```
Open an existing `.xlsx` file.

```json
{
  "type": "EXISTING",
  "path": "secured-workbook.xlsx",
  "security": {
    "password": "GridGrind-2026"
  }
}
```

Open an encrypted existing `.xlsx` package by supplying `source.security.password`.

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
{
  "type": "SAVE_AS",
  "path": "path/to/output.xlsx"
}
```
Write the workbook to the given path, creating parent directories as needed.

```json
{
  "type": "SAVE_AS",
  "path": "secured-output.xlsx",
  "security": {
    "encryption": {
      "password": "GridGrind-2026"
    },
    "signature": {
      "pkcs12Path": "signing-material.p12",
      "keystorePassword": "changeit",
      "keyPassword": "changeit",
      "alias": "gridgrind-signing"
    }
  }
}
```

`security.encryption` applies OOXML package encryption to the persisted workbook. Omit
`encryption.mode` to use the default `AGILE` mode.

`security.signature` applies OOXML package signing during persistence using a PKCS#12 keystore.
`pkcs12Path` must point to a readable `.p12` or `.pfx` file, and `keystorePassword` plus
`keyPassword` must unlock the selected key entry. Omit `alias` to use the sole keystore entry or
the first key entry POI can resolve.

Relative `path` values resolve from the current working directory.

The save path must end in `.xlsx`.

The response includes two path fields:
- `requestedPath` — the literal `path` string from the request.
- `executionPath` — the absolute normalized path where the file was actually written.

They are identical when an absolute path with no `..` segments is supplied. They differ when a
relative path (e.g. `"report.xlsx"`) or a path containing `..` segments is used.

```json
{
  "type": "OVERWRITE"
}
```
Overwrite the source file (requires `source.type=EXISTING`). `OVERWRITE` does not accept its own
`path` field; it always writes back to `source.path`. The response includes `sourcePath` (the
original source path string) and `executionPath` (the absolute normalized path).

```json
{
  "type": "OVERWRITE",
  "security": {
    "signature": {
      "pkcs12Path": "signing-material.p12",
      "keystorePassword": "changeit",
      "alias": "gridgrind-signing"
    }
  }
}
```

Use `OVERWRITE.security.signature` when persisting mutations to a signed source workbook. Unchanged
signed sources can be copied or overwritten without re-signing, but once a signed workbook is
mutated GridGrind requires explicit signature configuration before it will persist the result.

Omit `persistence` entirely or use `{ "type": "NONE" }` to run mutations, assertions, and
inspections without saving.

---

## Cell Values

Used in `SET_CELL`, `SET_RANGE`, and `APPEND_ROW`:

```json
{ "type": "TEXT",      "source": { "type": "INLINE", "text": "Origin" } }
{
  "type": "RICH_TEXT",
  "runs": [
    {
      "source": { "type": "INLINE", "text": "Q2 " },
      "font": { "fontName": "Aptos", "fontColor": "#44546A" }
    },
    {
      "source": { "type": "INLINE", "text": "Budget" },
      "font": { "bold": true, "fontColor": "#C00000" }
    }
  ]
}
{ "type": "NUMBER",    "number": 8.40                }
{ "type": "BOOLEAN",   "bool": true                  }
{ "type": "FORMULA",   "source": { "type": "INLINE", "text": "SUM(B2:B3)" } }  // leading = is accepted and stripped
{ "type": "DATE",      "date": "2026-03-25"           }
{ "type": "DATE_TIME", "dateTime": "2026-03-25T10:15:30" }
{ "type": "BLANK"                                     }
```

`TEXT`, `FORMULA`, and rich-text run payloads are source-backed: author inline text with
`{ "type": "INLINE", "text": "..." }`, or use `UTF8_FILE` / `STANDARD_INPUT` when the value
should come from the execution environment instead of the request body.
`TEXT` requires non-empty resolved text. Use `BLANK` when you want the cell itself to become
empty instead of storing a string value.
`RICH_TEXT` writes an ordered, non-empty `runs` list. Every run must have non-empty resolved text,
and the optional `font` object reuses the same font-field vocabulary as the nested style contract:
`bold`, `italic`, `fontName`, `fontHeight`, `fontColor`, `underline`, and `strikeout`.
`FORMULA` payloads are scalar only. Array-formula braces such as `{=SUM(A1:A2*B1:B2)}` are
rejected as `INVALID_FORMULA`. `LAMBDA` and `LET` are currently rejected as `INVALID_FORMULA`
because Apache POI cannot parse them. Other newer Excel constructs may fail the same way. Loaded
formulas that POI parses but cannot evaluate surface as `UNSUPPORTED_FORMULA`.

---

## Operations

### ENSURE_SHEET

Create a sheet if it does not already exist. Does nothing if the sheet exists.

All mutation operations (`SET_CELL`, `SET_RANGE`, `CLEAR_RANGE`, `APPLY_STYLE`,
`SET_DATA_VALIDATION`, `CLEAR_DATA_VALIDATIONS`, `SET_CONDITIONAL_FORMATTING`,
`CLEAR_CONDITIONAL_FORMATTING`, `SET_HYPERLINK`, `CLEAR_HYPERLINK`, `SET_COMMENT`,
`CLEAR_COMMENT`, `SET_PICTURE`, `SET_CHART`, `SET_SHAPE`, `SET_EMBEDDED_OBJECT`,
`SET_SIGNATURE_LINE`, `SET_DRAWING_OBJECT_ANCHOR`, `DELETE_DRAWING_OBJECT`, `SET_AUTOFILTER`,
`CLEAR_AUTOFILTER`, `APPEND_ROW`, `AUTO_SIZE_COLUMNS`) require the target sheet to already exist.
`SET_TABLE` and `SET_PIVOT_TABLE` also require the target `table.sheetName` or
`pivotTable.sheetName` to exist. Use `ENSURE_SHEET` before the first write to any sheet.

```json
{
  "stepId": "ensure-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "ENSURE_SHEET"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

---

### RENAME_SHEET

Rename an existing sheet. The source sheet must already exist. `newSheetName` must be a valid
Excel sheet name and must not conflict with another sheet name.

```json
{
  "stepId": "rename-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "Archive"
  },
  "action": {
    "type": "RENAME_SHEET",
    "newSheetName": "History"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `newSheetName` | Yes | New sheet name. Must be valid and unique. |

---

### DELETE_SHEET

Delete an existing sheet. A workbook must retain at least one sheet and at least one visible
sheet, so attempting to delete the last remaining sheet or the last visible sheet returns
`INVALID_REQUEST`.

```json
{
  "stepId": "delete-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "Scratch"
  },
  "action": {
    "type": "DELETE_SHEET"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

---

### MOVE_SHEET

Move an existing sheet to a zero-based workbook position. `targetIndex` is evaluated at the time
the operation runs, after all earlier operations in the same request.

```json
{
  "stepId": "move-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "History"
  },
  "action": {
    "type": "MOVE_SHEET",
    "targetIndex": 0
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "newSheetName": "Budget Review",
  "position": {
    "type": "AT_INDEX",
    "targetIndex": 1
  },
  "source": {
    "type": "BY_NAME",
    "name": "Budget"
  }
}
```

```json
{
  "type": "COPY_SHEET",
  "newSheetName": "Budget Snapshot",
  "source": {
    "type": "BY_NAME",
    "name": "Budget"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `source` | Yes | `BY_NAME` selector for the existing sheet to copy. |
| `newSheetName` | Yes | New unique destination sheet name. |
| `position` | No | Copy position. Omit to append at end. |

`position` variants:

- `{"type":"APPEND_AT_END"}`
- `{"type":"AT_INDEX","targetIndex":1}`

---

### SET_ACTIVE_SHEET

Set the active sheet. Hidden sheets cannot be activated. The active sheet is always selected.

```json
{
  "stepId": "set-active-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "Budget Review"
  },
  "action": {
    "type": "SET_ACTIVE_SHEET"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

---

### SET_SELECTED_SHEETS

Set the selected visible sheet set. Duplicate or unknown sheet names are rejected. Workbook
summary reads return `selectedSheetNames` in workbook order while preserving the chosen active
sheet as the primary selected tab.

```json
{
  "stepId": "set-selected-sheets",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Budget",
      "Budget Review"
    ]
  },
  "action": {
    "type": "SET_SELECTED_SHEETS"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAMES` selector for one or more existing visible sheets. |

---

### SET_SHEET_VISIBILITY

Set one sheet visibility state. A workbook must retain at least one visible sheet, so hiding the
last visible sheet is rejected.

```json
{
  "stepId": "set-sheet-visibility",
  "target": {
    "type": "BY_NAME",
    "name": "Archive"
  },
  "action": {
    "type": "SET_SHEET_VISIBILITY",
    "visibility": "VERY_HIDDEN"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `visibility` | Yes | `VISIBLE`, `HIDDEN`, or `VERY_HIDDEN`. |

---

### SET_SHEET_PROTECTION

Enable sheet protection with the exact supported lock flags and an optional password. When
`password` is provided, GridGrind writes a password hash into the sheet-protection XML while
preserving the explicit authored lock flags.

```json
{
  "stepId": "set-sheet-protection",
  "target": {
    "type": "BY_NAME",
    "name": "Budget Review"
  },
  "action": {
    "type": "SET_SHEET_PROTECTION",
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
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `protection` | Yes | Supported lock-flag payload. |
| `password` | No | Optional nonblank sheet-protection password. Omit it to protect the sheet without a stored password hash. |

---

### CLEAR_SHEET_PROTECTION

Disable sheet protection entirely and clear any stored password hash.

```json
{
  "stepId": "clear-sheet-protection",
  "target": {
    "type": "BY_NAME",
    "name": "Budget Review"
  },
  "action": {
    "type": "CLEAR_SHEET_PROTECTION"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

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
{
  "type": "CLEAR_WORKBOOK_PROTECTION"
}
```

No additional fields.

---

### IMPORT_CUSTOM_XML_MAPPING

Import XML content into one existing workbook custom-XML mapping. The step target is always
`WorkbookSelector.CURRENT`. GridGrind imports data into an existing mapping definition; it does
not author new workbook map definitions.

```json
{
  "stepId": "import-custom-xml",
  "target": {
    "type": "CURRENT"
  },
  "action": {
    "type": "IMPORT_CUSTOM_XML_MAPPING",
    "mapping": {
      "locator": {
        "mapId": 1,
        "name": "CORSO_mapping"
      },
      "xml": {
        "type": "UTF8_FILE",
        "path": "examples/custom-xml-assets/custom-xml-update.xml"
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Always `WorkbookSelector.CURRENT`. |
| `mapping` | Yes | Custom-XML import payload containing the existing mapping locator plus the XML content source. |

`mapping` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `locator.mapId` | No | Stable workbook map id. Supply `mapId`, `name`, or both. |
| `locator.name` | No | Workbook mapping name. Supply `mapId`, `name`, or both. |
| `xml` | Yes | XML content source. Supports `INLINE`, `UTF8_FILE`, and `STANDARD_INPUT`. |

The locator must resolve to exactly one existing mapping.

---

### MERGE_CELLS

Merge a rectangular A1-style range into one displayed region. The range must span at least two
cells. Repeating the same merge is a no-op, but overlapping a different merged region fails.

```json
{
  "stepId": "merge-cells",
  "target": {
    "type": "BY_RANGE",
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
    "type": "BY_RANGE",
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
| `heightPoints` | Yes | Positive Excel row height in points. Must be finite and > 0 and ≤ 1,638.35 (Excel storage limit: 32,767 twips). |

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

GridGrind rejects the edit when Apache POI would leave a moved table, sheet-owned autofilter, or
data validation stale (`LIM-016`).

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
    "type": "BY_NAME",
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
    "type": "BY_NAME",
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
    "type": "BY_NAME",
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
    "type": "BY_NAME",
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
- `sheetDefaults.defaultColumnWidth` must be a whole number greater than `0`.
- `sheetDefaults.defaultRowHeightPoints` must be finite and greater than `0`.
- Each `ignoredErrors` entry requires one A1-style rectangular `range` plus one or more distinct
  `errorTypes`.
- `ignoredErrors` ranges must be unique within the request.

### SET_PRINT_LAYOUT

Apply one authoritative supported print-layout state to a sheet. Omitted nested fields normalize
to default or cleared state.

```json
{
  "stepId": "set-print-layout",
  "target": {
    "type": "BY_NAME",
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
    "type": "BY_NAME",
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

### SET_CELL

Write a typed value to a single cell using A1 notation. Existing style, hyperlink, and comment
state on the targeted cell is preserved.

```json
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "B4"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "FORMULA",
      "source": {
        "type": "INLINE",
        "text": "SUM(B2:B3)"
      }
    }
  }
}
```

FORMULA note: a leading `=` is accepted and stripped automatically. `"=SUM(B2:B3)"` and
`"SUM(B2:B3)"` are equivalent. Scalar `FORMULA` values stay scalar only; use
`SET_ARRAY_FORMULA` for contiguous array-formula groups.

DATE / DATE_TIME note: the required Excel number format is applied without discarding any existing
fill, border, font, alignment, or wrap state already present on the cell.

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
  "stepId": "set-range",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Inventory",
    "range": "A1:C3"
  },
  "action": {
    "type": "SET_RANGE",
    "rows": [
      [
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Origin"
          }
        },
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Kilos"
          }
        },
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Cost/kg"
          }
        }
      ],
      [
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Ethiopia Yirgacheffe"
          }
        },
        {
          "type": "NUMBER",
          "number": 150
        },
        {
          "type": "NUMBER",
          "number": 8.4
        }
      ],
      [
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Colombia Huila"
          }
        },
        {
          "type": "NUMBER",
          "number": 200
        },
        {
          "type": "NUMBER",
          "number": 7.8
        }
      ]
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `range` | Yes | A1-notation rectangular range (e.g. `A1:C3`). |
| `rows` | Yes | 2-D array of typed cell values. Dimensions must match `range`. |

---

### SET_ARRAY_FORMULA

Author one contiguous single-cell or multi-cell array-formula group. The step target is a
rectangular `RangeSelector.BY_RANGE`; the stored array range is exactly the selector range.
Inline formula text may include or omit a leading `=` or `{=...}` wrapper and GridGrind
normalizes the stored formula text before handing it to POI.

```json
{
  "stepId": "set-array-formula",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Calc",
    "range": "D2:D4"
  },
  "action": {
    "type": "SET_ARRAY_FORMULA",
    "formula": {
      "source": {
        "type": "INLINE",
        "text": "{=B2:B4*C2:C4}"
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Range selector for the stored array-formula group. |
| `range` | Yes | Contiguous A1-notation array-formula range. |
| `formula` | Yes | Array-formula payload. Inline text may include or omit a leading `=` or `{=...}` wrapper. |

---

### CLEAR_ARRAY_FORMULA

Remove the stored array-formula group targeted by any member cell. The step target is a
`CellSelector.BY_ADDRESS`; if the addressed cell is not part of an array-formula group, the step
fails explicitly instead of silently doing nothing.

```json
{
  "stepId": "clear-array-formula",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Calc",
    "address": "D3"
  },
  "action": {
    "type": "CLEAR_ARRAY_FORMULA"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Cell selector targeting any member of the stored array-formula group. |
| `address` | Yes | One member cell of the stored array-formula group. |

---

### CLEAR_RANGE

Remove all values, styles, hyperlinks, and comments from a rectangular range. Cells that have
never been written are silently skipped; this operation never fails because a cell does not
physically exist.

```json
{
  "stepId": "clear-range",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Inventory",
    "range": "A2:C10"
  },
  "action": {
    "type": "CLEAR_RANGE"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `range` | Yes | Range to clear. |

---

### SET_HYPERLINK

Attach or replace a hyperlink on one cell. The cell is created if needed.

```json
{
  "stepId": "set-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "A1"
  },
  "action": {
    "type": "SET_HYPERLINK",
    "target": {
      "type": "URL",
      "target": "https://example.com/catalog"
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
{
  "stepId": "clear-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "A1"
  },
  "action": {
    "type": "CLEAR_HYPERLINK"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `address` | Yes | A1-notation cell address. |

---

### SET_COMMENT

Attach or replace one comment on one cell. The cell is created if needed. GridGrind always
requires the authoritative plain `text` plus `author`, and can also author ordered rich-text
`runs` and an explicit comment `anchor`.

```json
{
  "stepId": "set-comment",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "B4"
  },
  "action": {
    "type": "SET_COMMENT",
    "comment": {
      "text": {
        "type": "INLINE",
        "text": "Lead review scheduled."
      },
      "author": "GridGrind",
      "visible": false,
      "runs": [
        {
          "font": {
            "bold": true,
            "fontColorTheme": 4,
            "fontColorTint": -0.2
          },
          "source": {
            "type": "INLINE",
            "text": "Lead"
          }
        },
        {
          "source": {
            "type": "INLINE",
            "text": " "
          }
        },
        {
          "font": {
            "italic": true,
            "fontColorIndexed": 17
          },
          "source": {
            "type": "INLINE",
            "text": "review scheduled."
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
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `address` | Yes | A1-notation cell address. |
| `comment` | Yes | Comment payload. |

`comment.visible` defaults to `false` when omitted.

`comment` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `text` | Yes | Source-backed plain-text payload. Inline text uses `{ "type": "INLINE", "text": "..." }`; file-backed and standard-input variants are also supported. The resolved text must be nonblank. |
| `author` | Yes | Nonblank comment author. |
| `visible` | No | Whether the comment box is shown by default. Defaults to `false`. |
| `runs` | No | Ordered rich-text run list. When present, the concatenated run text must equal `comment.text` exactly. |
| `anchor` | No | Explicit comment-box bounds using zero-based `firstColumn`, `firstRow`, `lastColumn`, and `lastRow`. |

---

### CLEAR_COMMENT

Remove any comment attached to one cell. The cell value, style, and hyperlink are left unchanged.
No-op when the cell does not physically exist.

```json
{
  "stepId": "clear-comment",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "B4"
  },
  "action": {
    "type": "CLEAR_COMMENT"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
      "source": {
        "type": "INLINE_BASE64",
        "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
      }
    },
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 0,
        "rowIndex": 4,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 2,
        "rowIndex": 8,
        "dx": 0,
        "dy": 0
      },
      "behavior": "MOVE_AND_RESIZE"
    },
    "description": {
      "type": "INLINE",
      "text": "Queue preview"
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `picture` | Yes | Authoritative picture payload. |

`picture` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. Replaces any existing drawing object of the same name. |
| `image` | Yes | Binary picture payload with `format` plus base64 bytes. |
| `anchor` | Yes | Authored drawing anchor. The current public contract supports only `TWO_CELL`. |
| `description` | No | Optional nonblank descriptive text stored on the picture. |

`image.format` accepts `EMF`, `WMF`, `PICT`, `JPEG`, `PNG`, `DIB`, `GIF`, `TIFF`, `EPS`, `BMP`,
or `WPG`.

`anchor` currently uses:

```json
{
  "type": "TWO_CELL",
  "from": {
    "columnIndex": 0,
    "rowIndex": 4,
    "dx": 0,
    "dy": 0
  },
  "to": {
    "columnIndex": 2,
    "rowIndex": 8,
    "dx": 0,
    "dy": 0
  },
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
      "from": {
        "columnIndex": 3,
        "rowIndex": 4,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 5,
        "rowIndex": 7,
        "dx": 0,
        "dy": 0
      },
      "behavior": "MOVE_DONT_RESIZE"
    },
    "presetGeometryToken": "roundRect",
    "text": {
      "type": "INLINE",
      "text": "Queue"
    }
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
      "from": {
        "columnIndex": 3,
        "rowIndex": 8,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 6,
        "rowIndex": 9,
        "dx": 0,
        "dy": 0
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `shape` | Yes | Authoritative shape payload. |

`shape` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. |
| `kind` | Yes | `SIMPLE_SHAPE` or `CONNECTOR`. |
| `anchor` | Yes | Authored drawing anchor. The current public contract supports only `TWO_CELL`. |
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
    "previewImage": {
      "format": "PNG",
      "source": {
        "type": "INLINE_BASE64",
        "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
      }
    },
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 6,
        "rowIndex": 4,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 8,
        "rowIndex": 9,
        "dx": 0,
        "dy": 0
      }
    },
    "payload": {
      "type": "INLINE_BASE64",
      "base64Data": "R3JpZEdyaW5kIGVtYmVkZGVkIHBheWxvYWQ="
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
| `anchor` | Yes | Authored drawing anchor. The current public contract supports only `TWO_CELL`. |

---

### SET_CHART

Create or mutate one named supported chart on one sheet. The public chart contract now models one
chart with ordered `plots`, so authored requests can build one supported plot or one multi-plot
combo chart. Supported authored plot families are `AREA`, `AREA_3D`, `BAR`, `BAR_3D`,
`DOUGHNUT`, `LINE`, `LINE_3D`, `PIE`, `PIE_3D`, `RADAR`, `SCATTER`, `SURFACE`, and
`SURFACE_3D`. Series bind to workbook ranges, defined names, or literal values depending on the
plot. Existing unsupported loaded detail is preserved on unrelated edits and rejected for
authoritative mutation.

```json
{
  "type": "SET_CHART",
  "sheetName": "Ops",
  "chart": {
    "name": "OpsChart",
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 4,
        "rowIndex": 0,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 8,
        "rowIndex": 12,
        "dx": 0,
        "dy": 0
      },
      "behavior": "MOVE_AND_RESIZE"
    },
    "title": {
      "type": "TEXT",
      "source": {
        "type": "INLINE",
        "text": "Roadmap"
      }
    },
    "legend": {
      "type": "VISIBLE",
      "position": "TOP_RIGHT"
    },
    "displayBlanksAs": "SPAN",
    "plotOnlyVisibleCells": false,
    "plots": [
      {
        "type": "BAR",
        "varyColors": true,
        "barDirection": "COLUMN",
        "series": [
          {
            "title": {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Plan"
              }
            },
            "categories": {
              "type": "REFERENCE",
              "formula": "ChartCategories"
            },
            "values": {
              "type": "REFERENCE",
              "formula": "Ops!$B$2:$B$4"
            }
          },
          {
            "title": {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Actual"
              }
            },
            "categories": {
              "type": "REFERENCE",
              "formula": "ChartCategories"
            },
            "values": {
              "type": "REFERENCE",
              "formula": "ChartActual"
            }
          }
        ]
      }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `chart` | Yes | Authoritative supported-chart payload. |

`chart` common fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local chart name. Reuses the existing chart when the name already exists. |
| `anchor` | Yes | Authored chart-frame anchor. Currently only `TWO_CELL` is supported. |
| `title` | No | `NONE`, `TEXT`, or `FORMULA`. Defaults to `NONE`. |
| `legend` | No | `HIDDEN` or `VISIBLE`. `VISIBLE.position` defaults to `RIGHT` when omitted. |
| `displayBlanksAs` | No | `GAP`, `SPAN`, or `ZERO`. Defaults to `GAP`. |
| `plotOnlyVisibleCells` | No | Whether hidden cells are ignored. Defaults to `true`. |
| `plots` | Yes | Ordered non-empty plot list. Use one plot for a simple chart or multiple plots for a combo chart. |

`plot` common fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `type` | Yes | `AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`, `LINE_3D`, `PIE`, `PIE_3D`, `RADAR`, `SCATTER`, `SURFACE`, or `SURFACE_3D`. |
| `varyColors` | No | Whether Excel varies series colors automatically. Defaults to `false`. |
| `series` | Yes | Ordered non-empty series list. |

Axis-backed plot families (`AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `LINE`, `LINE_3D`, `RADAR`,
`SCATTER`, `SURFACE`, `SURFACE_3D`) may also provide explicit ordered `axes`. When omitted,
GridGrind writes the default axis set for that plot family.

`series` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `title` | No | `NONE`, `TEXT`, or `FORMULA`. Defaults to `NONE`. |
| `categories` | Yes | `REFERENCE`, `STRING_LITERAL`, or `NUMERIC_LITERAL` data source for the category axis. |
| `values` | Yes | `REFERENCE`, `STRING_LITERAL`, or `NUMERIC_LITERAL` data source for the value axis. |
| `smooth` | No | Line and scatter smoothing flag when the target plot family supports it. |
| `markerStyle` | No | Marker shape for line and scatter series. |
| `markerSize` | No | Marker size between `2` and `72`. |
| `explosion` | No | Pie or doughnut slice explosion amount. Must not be negative. |

Variant-specific fields:

| Chart type | Field | Required | Description |
|:-----------|:------|:---------|:------------|
| `AREA`, `AREA_3D`, `LINE`, `LINE_3D` | `grouping` | No | Plot grouping. Defaults to `STANDARD`. |
| `AREA_3D`, `LINE_3D`, `BAR_3D` | `gapDepth` | No | Optional 3D gap depth. |
| `BAR`, `BAR_3D` | `barDirection` | No | `COLUMN` or `BAR`. Defaults to `COLUMN`. |
| `BAR`, `BAR_3D` | `grouping` | No | Bar grouping. Defaults to `CLUSTERED`. |
| `BAR` | `gapWidth`, `overlap` | No | Optional bar spacing overrides. `overlap` must be between `-100` and `100`. |
| `BAR_3D` | `gapWidth`, `shape` | No | Optional bar spacing and 3D bar shape overrides. |
| `DOUGHNUT` | `firstSliceAngle`, `holeSize` | No | `firstSliceAngle` must be between `0` and `360`; `holeSize` must be between `10` and `90`. |
| `PIE` | `firstSliceAngle` | No | Integer between `0` and `360`. |
| `RADAR` | `style` | No | Radar style. Defaults to `STANDARD`. |
| `SCATTER` | `style` | No | Scatter style. Defaults to `LINE_MARKER`. |
| `SURFACE`, `SURFACE_3D` | `wireframe` | No | Whether the surface plot is written as wireframe. Defaults to `false`. |

Reference-backed `categories` and `values` accept either contiguous A1-style references such as
`Ops!$B$2:$B$4` or defined names such as `ChartActual`.
Formula-backed chart titles and series titles must resolve to one cell, either directly or
through a defined name that resolves to one cell.
Validation failures are non-mutating: GridGrind rejects invalid authored chart payloads without
creating a partial chart or partially mutating an existing supported chart.

---

### SET_SIGNATURE_LINE

Create or replace one named signature line on one sheet. Signature lines are VML-backed drawing
objects with authored signer suggestions, comment permissions, optional caption text, and an
optional plain-signature preview image.

```json
{
  "type": "SET_SIGNATURE_LINE",
  "sheetName": "Approvals",
  "signatureLine": {
    "name": "BudgetSignature",
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 1,
        "rowIndex": 1,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 4,
        "rowIndex": 6,
        "dx": 0,
        "dy": 0
      },
      "behavior": "MOVE_AND_RESIZE"
    },
    "allowComments": false,
    "signingInstructions": "Review the budget before signing.",
    "suggestedSigner": "Ada Lovelace",
    "suggestedSigner2": "Finance",
    "suggestedSignerEmail": "ada@example.com",
    "invalidStamp": "invalid",
    "plainSignature": {
      "format": "PNG",
      "source": {
        "type": "INLINE_BASE64",
        "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target sheet. |
| `signatureLine` | Yes | Authoritative signature-line payload. |

`signatureLine` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. Replaces any existing signature line of the same name. |
| `anchor` | Yes | Authored drawing anchor. The current public contract supports only `TWO_CELL`. |
| `allowComments` | No | Whether Excel allows comments alongside the signature line. Defaults to `true`. |
| `signingInstructions` | No | Optional nonblank instruction text shown in Excel's signing UI. |
| `suggestedSigner` | No | Optional nonblank signer name. |
| `suggestedSigner2` | No | Optional nonblank second signer line, often used for a title or department. |
| `suggestedSignerEmail` | No | Optional nonblank signer email address. |
| `caption` | No | Optional nonblank caption text limited to at most three lines. |
| `invalidStamp` | No | Optional nonblank stamp text used for invalid-signature rendering. |
| `plainSignature` | No | Optional preview image payload for the plain signature. Reuses the same `format` plus binary `source` shape as `SET_PICTURE.picture.image`. |

At least one signer-suggestion field (`caption`, `suggestedSigner`, `suggestedSigner2`, or
`suggestedSignerEmail`) must be present. Validation failures are non-mutating: invalid
signature-line payloads do not delete an existing signature line or leave a partial VML shape
behind.

---

### SET_DRAWING_OBJECT_ANCHOR

Move one existing named drawing object by replacing its anchor authoritatively. Supported named
targets include pictures, signature lines, simple shapes, connectors, embedded objects, and
supported charts.

```json
{
  "stepId": "set-drawing-object-anchor",
  "target": {
    "type": "BY_NAME",
    "sheetName": "Ops",
    "objectName": "OpsPicture"
  },
  "action": {
    "type": "SET_DRAWING_OBJECT_ANCHOR",
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 1,
        "rowIndex": 5,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 3,
        "rowIndex": 9,
        "dx": 0,
        "dy": 0
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAME` drawing-object selector for the existing sheet-local object. |
| `anchor` | Yes | Replacement authored drawing anchor. The current public contract supports only `TWO_CELL`. |

---

### DELETE_DRAWING_OBJECT

Delete one existing named drawing object from one sheet.

```json
{
  "stepId": "delete-drawing-object",
  "target": {
    "type": "BY_NAME",
    "sheetName": "Ops",
    "objectName": "OpsConnector"
  },
  "action": {
    "type": "DELETE_DRAWING_OBJECT"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAME` drawing-object selector for the existing sheet-local object to delete. Signature-line deletes also remove the associated preview-image relation when one exists. |

---

### APPLY_STYLE

Apply a style patch to a rectangular range. Only fields present in `style` are changed;
unspecified style properties are left untouched.

```json
{
  "stepId": "apply-style",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Inventory",
    "range": "A1:C1"
  },
  "action": {
    "type": "APPLY_STYLE",
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
        "fontHeight": {
          "type": "POINTS",
          "points": 13
        },
        "fontColor": "#FFFFFF"
      },
      "fill": {
        "pattern": "THIN_HORIZONTAL_BANDS",
        "foregroundColor": "#1F4E78",
        "backgroundColor": "#D9E2F3"
      },
      "border": {
        "all": {
          "style": "THIN",
          "color": "#FFFFFF"
        },
        "bottom": {
          "style": "DOUBLE",
          "color": "#D9E2F3"
        }
      },
      "protection": {
        "locked": true
      }
    }
  }
}
```

```json
{
  "stepId": "apply-style",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Inventory",
    "range": "J2:J3"
  },
  "action": {
    "type": "APPLY_STYLE",
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
            {
              "position": 0.0,
              "color": {
                "rgb": "#1F497D"
              }
            },
            {
              "position": 1.0,
              "color": {
                "theme": 4,
                "tint": 0.45
              }
            }
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
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
- `fill.gradient.type="LINEAR"` accepts `degree` only. `fill.gradient.type="PATH"` accepts
  `left`, `right`, `top`, and `bottom` only. Mixing the two geometry models is rejected as an
  invalid request.

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
  "stepId": "set-data-validation",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Inventory",
    "range": "B2:B200"
  },
  "action": {
    "type": "SET_DATA_VALIDATION",
    "validation": {
      "rule": {
        "type": "EXPLICIT_LIST",
        "values": [
          "Queued",
          "Done",
          "Blocked"
        ]
      },
      "allowBlank": true,
      "prompt": {
        "title": {
          "type": "INLINE",
          "text": "Status"
        },
        "text": {
          "type": "INLINE",
          "text": "Pick one workflow state."
        }
      },
      "errorAlert": {
        "style": "STOP",
        "title": {
          "type": "INLINE",
          "text": "Invalid status"
        },
        "text": {
          "type": "INLINE",
          "text": "Use one of the allowed values."
        }
      }
    }
  }
}
```

```json
{
  "stepId": "set-data-validation",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Inventory",
    "range": "C2:C200"
  },
  "action": {
    "type": "SET_DATA_VALIDATION",
    "validation": {
      "rule": {
        "type": "WHOLE_NUMBER",
        "operator": "BETWEEN",
        "formula1": "1",
        "formula2": "5"
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
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
| `prompt` | Optional prompt box with source-backed `title` and `text` values plus optional `showPromptBox` (defaults to `true`). |
| `errorAlert` | Optional error box with `style`, source-backed `title` / `text`, and optional `showErrorBox` (defaults to `true`). |

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
  "stepId": "clear-data-validations",
  "target": {
    "type": "ALL_ON_SHEET",
    "sheetName": "Inventory"
  },
  "action": {
    "type": "CLEAR_DATA_VALIDATIONS"
  }
}
```

```json
{
  "stepId": "clear-data-validations",
  "target": {
    "type": "BY_RANGES",
    "sheetName": "Inventory",
    "ranges": [
      "B4",
      "C8:D10"
    ]
  },
  "action": {
    "type": "CLEAR_DATA_VALIDATIONS"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `target` | Yes | `ALL_ON_SHEET` or `BY_RANGES` selector for the target ranges. |

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
  "stepId": "set-conditional-formatting",
  "target": {
    "type": "BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_CONDITIONAL_FORMATTING",
    "conditionalFormatting": {
      "ranges": [
        "D2:D200"
      ],
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
}
```

```json
{
  "stepId": "set-conditional-formatting",
  "target": {
    "type": "BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_CONDITIONAL_FORMATTING",
    "conditionalFormatting": {
      "ranges": [
        "K2:K200"
      ],
      "rules": [
        {
          "type": "COLOR_SCALE_RULE",
          "stopIfTrue": false,
          "thresholds": [
            {
              "type": "MIN"
            },
            {
              "type": "PERCENTILE",
              "value": 50.0
            },
            {
              "type": "MAX"
            }
          ],
          "colors": [
            {
              "rgb": "#AA2211"
            },
            {
              "rgb": "#FFDD55"
            },
            {
              "rgb": "#11CC66"
            }
          ]
        }
      ]
    }
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

Removes conditional-formatting blocks on the sheet that intersect the provided range selector.
`ALL_ON_SHEET` clears the sheet. `BY_RANGES` removes every stored block whose target ranges
intersect any of the selected A1 ranges.

```json
{
  "stepId": "clear-conditional-formatting",
  "target": {
    "type": "ALL_ON_SHEET",
    "sheetName": "Inventory"
  },
  "action": {
    "type": "CLEAR_CONDITIONAL_FORMATTING"
  }
}
```

```json
{
  "stepId": "clear-conditional-formatting",
  "target": {
    "type": "BY_RANGES",
    "sheetName": "Inventory",
    "ranges": [
      "D2:D200",
      "F2:F20"
    ]
  },
  "action": {
    "type": "CLEAR_CONDITIONAL_FORMATTING"
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
  "stepId": "set-autofilter",
  "target": {
    "type": "BY_RANGE",
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
    "type": "BY_NAME",
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
    "type": "BY_NAME",
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

---

### AUTO_SIZE_COLUMNS

Resize columns to fit their content. Applies to all columns with data on the sheet.

```json
{
  "stepId": "auto-size-columns",
  "target": {
    "type": "BY_NAME",
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
