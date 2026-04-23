---
afad: "3.5"
version: "0.56.0"
domain: REQUEST_EXECUTION_REFERENCE
updated: "2026-04-23"
route:
  keywords: [gridgrind, request, source, persistence, execution, formula-environment, source-backed, input, calculation, journal, event-read, streaming-write]
  questions: ["what does a gridgrind request look like", "how do source-backed inputs work in gridgrind", "how does execution.calculation work", "what is the response journal", "how do event read and streaming write work"]
---

# Request And Execution Reference

**Purpose**: Canonical request-envelope reference for GridGrind `.xlsx` workflows: source,
persistence, source-backed authored inputs, execution policy, formula-environment bindings, and
core value shapes.
**Companion references**: [OPERATIONS.md](./OPERATIONS.md),
[WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md),
[CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md),
[STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)
**Limits**: See [LIMITATIONS.md](./LIMITATIONS.md) for hard ceilings and mode restrictions.

The long-form step reference is intentionally split. This document owns the request envelope and
execution policy only; the detailed mutation, assertion, and inspection sections live in the
focused references linked above. The Java authoring layer emits this same envelope through
`GridGrindPlan.toPlan()`, `toJsonBytes()`, and `toJsonString()`. See
[JAVA_AUTHORING.md](./JAVA_AUTHORING.md) when you want to build the request from Java instead of
hand-writing JSON.

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

- When the CLI reads a request via `--request <path>`, relative `UTF8_FILE` and `FILE` paths
  resolve from that request file's directory. Without `--request`, they resolve in the current
  execution environment.
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

When the CLI reads the request from `--request <path>`, relative request-owned paths inside the
JSON follow the request file directory. That includes `source.path`, `persistence.path`,
source-backed `UTF8_FILE` / `FILE` payloads, `formulaEnvironment.externalWorkbooks[*].path`, and
`persistence.security.signature.pkcs12Path`. The CLI flags themselves are separate: `--request`
and `--response` still resolve from the shell working directory.

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
| `externalWorkbooks` | No | Workbook-name to path bindings used to satisfy formulas such as `[rates.xlsx]Sheet1!A1`. Each `path` follows the same request-owned path rule described above. |
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
- `EVALUATE_TARGETS` addresses must point at existing formula cells. A missing physical cell can
  surface `CELL_NOT_FOUND`; an existing non-formula cell is rejected as `INVALID_REQUEST`.
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

When the CLI reads the request via `--request <path>`, relative `path` values resolve from that
request file's directory. Without `--request`, they resolve in the current execution environment.

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
the first key entry POI can resolve. `pkcs12Path` follows the same request-owned path rule as
other request file paths.

When the CLI reads the request via `--request <path>`, relative persistence `path` values resolve
from that request file's directory. Without `--request`, they resolve in the current execution
environment.

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
