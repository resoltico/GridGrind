---
afad: "5.0.1"
version: "0.74.0"
domain: REQUEST_EXECUTION_REFERENCE
updated: "2026-08-27"
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
  resolve from that request file's directory. When the request JSON arrives on stdin, pass
  `--execution-root <path>` and those same relative paths resolve from that explicit directory.
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

## Doctor Requests

`--doctor-request` validates request shape, each bound operation's target contract,
execution-mode rules, source-backed authored input resolution, and existing workbook-source
accessibility without mutating a workbook.

- It resolves `UTF8_FILE`, `FILE`, and `STANDARD_INPUT` authored payloads early, so missing or
  unreadable authored inputs can fail under `journal.inputResolution`. Each source-resolution
  problem carries `context.json.jsonPath` for the exact authored leaf, such as
  `steps[0].action.rows.cells[0][1].source.path`; raw requests retain that leaf's UTF-8 byte
  offset as well.
- It also preflights `source.type: EXISTING` workbook access, so missing or unreadable
  `source.path` workbooks can already fail during doctoring under `OPEN_WORKBOOK`.
- It collects every independently observable request-intake defect in one pass, including invalid UTF-8, duplicate keys, unknown fields, omitted required fields, explicit nulls, malformed scalar values, missing or unknown type discriminators, and constructor-level field validation failures. Valid sibling fragments remain available for their own operation-contract checks; a rule whose own prerequisite fragment is malformed is suppressed rather than guessed.
- When every request fragment binds, doctoring batches independent source-backed input and existing-workbook preflight failures even when the plan also has static operation or execution-policy failures. This includes independent authored inputs nested in the same step, so one unreadable cell value does not hide a sibling value's failure. A source-backed input failure does not hide an inaccessible existing workbook, and doctoring never mutates or persists a workbook.
- Normal execution performs the same request intake before any workbook work begins. For the
  same request bytes, a rejected `--request` command emits the same ordered problem core in
  `CommandError.problems` that `--doctor-request` returns in `RequestDoctorReport.problems` for
  request-intake and static findings. Execution rejects static findings before workbook access;
  once static validation passes, a failed source/input preflight completes zero steps and persists
  no workbook.
- It emits its own machine-readable `RequestDoctorReport`; a report with findings is
  `valid:false`, not a rejected command result.
- When the request JSON arrives on stdin, pass `--execution-root <path>` so doctoring uses one
  explicit request root instead of ambient process state.
- `--response <path>` works here too, so the doctor report can be captured to a file instead of
  stdout when the workflow needs a saved artifact.
- Without `--response`, each command writes exactly one primary JSON payload to stdout. With it,
  the already-rendered payload goes only to the requested file. If that file cannot be written,
  GridGrind recovers those unchanged bytes to stdout when stdout is writable and writes one
  transport-only JSON notice to stderr. GridGrind never moves a primary payload to stderr.
- For `execute`, a requested response file is reserved with create-new and no-follow semantics
  after request validation but before input binding or workbook work. An existing, directory, or
  unwritable response path prevents execution; stderr receives a typed `CliTransportNotice` with
  `reason`, `wroteTo=NOT_DELIVERED`, and `responsePath`. A reservation stays open until its one
  rendered result is written, so GridGrind never reopens the path by name or overwrites a prior
  response file.

---

## Request Structure

```json
{
  "protocolVersion": "V2",
  "source":      { ... },
  "persistence": { ... },
  "steps": [ ... ]
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `protocolVersion` | Yes | Wire-contract version. The current public value is `V2`. |
| `source` | Yes | Where the workbook comes from. |
| `persistence` | Yes | Where and whether to save. Use `{"type":"NONE"}` for unsaved runs. |
| `execution` | No | Optional execution policy for low-memory mode selection, structured journaling, and formula calculation handling. Omit it for the standard full-XSSF path with `SUMMARY` journaling and `DO_NOT_CALCULATE`, or supply it explicitly when you need non-default behavior. |
| `formulaEnvironment` | No | Optional evaluator configuration for external workbook bindings, missing-workbook policy, and template-backed UDF toolpacks. Omit it when the default evaluator is intended, or supply it when execution needs workbook bindings, `USE_CACHED_VALUE`, or UDF toolpacks. |
| `steps` | Yes | Ordered list of workbook mutations, assertions, and inspections. Send `[]` for a no-op plan. Every non-empty step needs a caller-defined `stepId`; `stepId` values must be unique within `steps[]` and must match `[A-Za-z0-9._-]+`. |

Every tagged request union uses `type` as its discriminator field: `source`, `persistence`,
`action`, `query`, cell values, hyperlink targets, selectors, and named-range scopes.
Every step object carries a caller-defined `stepId` plus exactly one of `action`, `assertion`, or
`query`. `stepId` values must be unique within `steps[]` and must match `[A-Za-z0-9._-]+`. Step
kind is inferred from that field; request steps do not carry a separate `step.type`.
`gridgrind --print-request-template` emits the canonical minimal valid request with the minimal
top-level envelope shown above. Add `execution` and/or `formulaEnvironment` only when the request
needs non-default behavior.

When the CLI reads the request from `--request <path>`, relative request-owned paths inside the
JSON follow the request file directory. That includes `source.path`, `persistence.path`,
source-backed `UTF8_FILE` / `FILE` payloads, `formulaEnvironment.externalWorkbooks[*].path` when
present, and `persistence.security.signature.signature.pkcs12Path`. The CLI flags themselves are separate:
`--request` and `--response` still resolve from the shell working directory, as do
`--execution-root` and `--temp-root`. Execution scratch is not request-rooted: without `--temp-root`, GridGrind
creates one private per-run scratch directory under the OS temporary-file root; with
`--temp-root <path>`, it creates that private scratch directory under the supplied parent path,
and best-effort cleanup removes it on normal command completion. Encrypted OOXML plaintext temp
workbooks always stay in private OS temp rather than the request root, execution root, or
`--temp-root` parent.

### Formula Environment

`formulaEnvironment` is optional at the top level. Omit it when the default evaluator is
intended. Supply it when server-side formula evaluation needs external workbook bindings,
cached-value fallback for unresolved external references, or template-backed UDFs. When the block
is present, the nested fields below stay explicit on the wire.

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
| `externalWorkbooks` | No | Workbook-name to path bindings used to satisfy formulas such as `[rates.xlsx]Sheet1!A1`. Each `path` follows the same request-owned path rule described above. Defaults to `[]` when omitted. |
| `missingWorkbookPolicy` | No | `ERROR` or `USE_CACHED_VALUE`. Defaults to `ERROR` when omitted. |
| `udfToolpacks` | No | Named collections of template-backed UDFs. Defaults to `[]` when omitted. |

For `udfToolpacks.functions`, `maximumArgumentCount` is optional and defaults to
`minimumArgumentCount`. `formulaTemplate` may reference `ARG1`, `ARG2`, and higher placeholders.

### Execution Policy

`execution` is optional at the top level. Omit it when the standard `FULL_XSSF` / `SUMMARY` /
`DO_NOT_CALCULATE` policy is intended. Supply it when the request needs a non-default execution
mode, journal level, calculation policy, or assertion policy. When the block is present, each nested field may still
be omitted to keep its own default, so callers can send only the execution axis they want to
override.

```json
{
  "execution": {
    "journal": {
      "level": "VERBOSE"
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `mode` | No | Explicit execution-mode variant selection through `type=FULL_XSSF`, `EVENT_READ`, or `STREAMING_WRITE`. Defaults to `FULL_XSSF` when omitted. |
| `journal` | No | Explicit structured-journal policy. Defaults to `SUMMARY` when omitted. |
| `calculation` | No | Explicit formula-calculation policy covering immediate evaluation, cache clearing, and workbook-open recalc flags. Defaults to `DO_NOT_CALCULATE` with `markRecalculateOnOpen=false` when omitted. |
| `assertionMode` | No | `FAIL_FAST` stops at the first failed assertion. `COLLECT` evaluates every assertion in the terminal assertion phase before returning the first canonical failure. Defaults to `FAIL_FAST`. |

- `execution.mode.type: EVENT_READ` selects the low-memory XSSF event-model reader. It supports only
  `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY` (`LIM-019`).
- `execution.mode.type: STREAMING_WRITE` selects the low-memory SXSSF writer. It requires
  `source.type: NEW`, supports only `ENSURE_SHEET` and `APPEND_ROW`,
  requires `execution.calculation.strategy=DO_NOT_CALCULATE`,
  allows `markRecalculateOnOpen=true`, and
  requires at least one `ENSURE_SHEET` mutation (`LIM-020`). GridGrind keeps shared strings
  enabled in this mode so large repeated-text workbooks do not balloon into inline-string-heavy
  OOXML packages.
- `execution.mode.type: FULL_XSSF` is the default full workbook read/write path with no low-memory restrictions.
- `execution.journal.level` accepts `SUMMARY`, `NORMAL`, and `VERBOSE`.
- `SUMMARY` is the default. It keeps the response stable by omitting phase timestamps, using
  `durationMillis=0`, recording compact resolved-target summaries, and suppressing live progress
  output.
- `NORMAL` keeps the structured response journal and adds expanded resolved-target summaries plus
  observational timing telemetry.
- `execution.calculation.strategy` accepts `DO_NOT_CALCULATE`, `DEFERRED_CALCULATION`,
  `EVALUATE_ALL`, `EVALUATE_TARGETS`, `REQUIRE_EVALUATION`, and `CLEAR_CACHES_ONLY`.
  `DEFERRED_CALCULATION` reports capability warnings without attempting server-side evaluation.
  The two evaluation strategies are lenient: unevaluable formulas remain unchanged, calculation
  reports `PARTIAL`, and each affected formula emits `FORMULA_NOT_EVALUATED`.
  `REQUIRE_EVALUATION` instead fails when any formula cannot be evaluated immediately.
- `execution.calculation.markRecalculateOnOpen` persists Excel's workbook-level recalc-on-open
  flag without requiring an extra mutation step.
- `execution.assertionMode=COLLECT` makes the first assertion step the start of a terminal
  verification phase. No later `MUTATION` step is legal; `INSPECTION` steps may interleave, and
  every assertion is returned as `PASSED` or `FAILED` before the run returns its canonical first
  assertion failure.
- `EVALUATE_TARGETS` addresses must point at existing formula cells. A missing physical cell can
  surface `CELL_NOT_FOUND`; an existing non-formula cell is rejected as `INVALID_REQUEST`.
- `VERBOSE` keeps the `NORMAL` response journal detail and streams fine-grained progress as compact
  JSONL on stderr. Each line is an `ExecutionProgressEvent`; `--pretty` never indents those lines.
- `EVENT_READ` can run directly against an existing workbook when the request is read-only and
  unsaved. If the request also performs full-XSSF mutations, GridGrind materializes the mutated
  workbook state and then performs the summary reads through the event model.
- Execution mode is one closed variant, not a read/write cross-product. Choose exactly one of
  `FULL_XSSF`, `EVENT_READ`, or `STREAMING_WRITE` for the whole request.

### Execution Results And Journal

Only `--request` emits `WorkbookResult` after workbook execution begins. Every `SUCCEEDED` and
`FAILED` execution result includes one top-level `persistence` outcome, structured `journal`,
`warnings`, `assertions`, and `inspections`; only `FAILED` carries the singular top-level `problem`.
Persist-workbook failures keep the save outcome at the response root; their `problem.context` adds
`sourceWorkbookPath` or `persistencePath` directly instead of nesting a second `persistence`
object. The excerpt below focuses on that canonical persistence outcome and its journal telemetry:

```json
{
  "status": "SUCCEEDED",
  "protocolVersion": "V2",
  "planId": "budget-pass",
  "persistence": {
    "type": "SAVE_AS",
    "requestedPath": "out/budget-reviewed.xlsx",
    "write": {
      "status": "WRITTEN",
      "executionPath": "/work/out/budget-reviewed.xlsx"
    }
  },
  "journal": {
    "level": "NORMAL",
    "source": {
      "type": "EXISTING",
      "path": "budget.xlsx"
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
    "outcome": {
      "status": "SUCCEEDED",
      "plannedStepCount": 1,
      "completedStepCount": 1,
      "durationMillis": 29
    }
  },
  "warnings": []
}
```

| Field | Description |
|:------|:------------|
| `planId` | Caller-supplied plan correlation ID when present. It belongs to the top-level `WorkbookResult`, not journal telemetry. |
| `level` | `SUMMARY`, `NORMAL`, or `VERBOSE`, matching the effective `execution.journal.level`. |
| `source` | Structured echo of the authored workbook source family and, when applicable, the authored source path string. |
| `validation`, `inputResolution`, `open`, `persistencePhase`, `close` | Top-level pipeline phase summaries. `SUMMARY` keeps these phases timestamp-free with `durationMillis=0`; `NORMAL` and `VERBOSE` add observational `startedAt`, `finishedAt`, and non-zero timing where applicable. `NOT_STARTED` and `NOT_REQUESTED` always omit timestamps and use `durationMillis=0`. `inputResolution` records source-backed file/stdin loading before workbook open. |
| `calculation` | Top-level calculation telemetry. `preflight` classifies authored formulas and `execution` records the requested evaluation or cache-clearing work. |
| `steps[]` | Ordered per-step telemetry including `resolvedTargets`, phase timing, outcome, and optional failure classification. `resolvedTargets` is compact in `SUMMARY` and expanded in `NORMAL`/`VERBOSE`. |
| `outcome` | Whole-run status plus `plannedStepCount`, `completedStepCount`, total `durationMillis`, and optional `failedStepIndex`, `failedStepId`, and `failureCode` when the run failed. |

`VERBOSE` keeps the full response journal and emits one compact `ExecutionProgressEvent` JSON line
to stderr for each lifecycle transition. The event carries `timestamp`, `category`, `status`, optional
`problemCode`, and optional `stepIndex`/`stepId`; it never carries an unstructured `detail` string.
When response-file fallback also occurs, its structured transport notice is one additional stderr JSON line.

The top-level response `persistence` field is the canonical save outcome. The journal records
`persistencePhase` timing only; it no longer repeats the intended or actual save target, and
`problem.context` uses `sourceWorkbookPath` / `persistencePath` fields instead of another nested
`persistence` block.

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
  "protocolVersion": "V2",
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
        "type": "WORKBOOK_CURRENT"
      },
      "query": {
        "type": "ANALYZE_WORKBOOK_FINDINGS"
      }
    }
  ]
}
```

Successful responses may include a `warnings` array. Warning locations are typed: `STEP` identifies
an authored step, `REQUEST_PATH` identifies a request-owned file path, `REQUEST_BYTE_OFFSET`
identifies an exact request-stream position, and `FORMULA_CELL` identifies one formula. Current
warnings flag same-request sheet names with spaces referenced in formulas without single quotes,
contained absolute request paths, and a leading UTF-8 BOM (`UTF8_BOM_IGNORED` at byte offset zero).
Use `'Sheet Name'!A1` syntax for the former; prefer relative paths for the latter so the request
remains portable across execution roots.

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
request file's directory. When the request JSON arrives on stdin, pass
`--execution-root <path>` and those same relative paths resolve from that explicit directory.

GridGrind supports `.xlsx` only. Paths ending in `.xls`, `.xlsm`, `.xlsb`, or any other
non-`.xlsx` extension are rejected as invalid requests.

---

## Persistence

The response `persistence.type` field always echoes the request `persistence.type` value, making
it straightforward to correlate request and response: a `SAVE_AS` request yields a `SAVE_AS`
response, an `OVERWRITE` request yields an `OVERWRITE` response, and a `NONE` request yields a
`NONE` response. When an `OVERWRITE` request fails before any `EXISTING` source path exists, the
response still reports `type=OVERWRITE` but omits `sourcePath` rather than inventing one.

```json
{
  "type": "SAVE_AS",
  "path": "path/to/output.xlsx",
  "ifExists": "REJECT"
}
```
Write the workbook to the given path. The destination parent directory must already exist so
GridGrind can bind it through a no-follow filesystem handle before execution begins.
`SAVE_AS.ifExists` is required: use `REJECT` to fail when the destination already exists, or
`REPLACE` to allow create-or-replace writes.

```json
{
  "type": "SAVE_AS",
  "path": "secured-output.xlsx",
  "ifExists": "REPLACE",
  "security": {
    "encryption": {
      "type": "ENCRYPT",
      "encryption": {
        "password": "GridGrind-2026",
        "cipher": "AES_256",
        "hash": "SHA_512"
      }
    },
    "signature": {
      "type": "SIGN",
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

`security.encryption` is an explicit policy: `NONE` deliberately writes plaintext, `ENCRYPT`
applies the nested OOXML write envelope, and `PRESERVE_SOURCE` reapplies a verified,
write-compatible encrypted source envelope. `PRESERVE_SOURCE` rejects plaintext sources during
preflight. GridGrind writes AGILE packages only; `mode` is not part of the request shape.
`cipher` defaults to `AES_256`, `hash` defaults to `SHA_512`, supported ciphers are
`AES_256` and `AES_192`, and supported hashes are `SHA_512`, `SHA_384`, and `SHA_256`.
Legacy STANDARD packages remain readable on inspection but are not authorable.

`security.signature` is an explicit policy: `NONE` deliberately writes an unsigned package and
removes any source package signatures; `SIGN` removes any source package signatures and applies
one fresh nested PKCS#12 signature during persistence.
`pkcs12Path` must point to a readable `.p12` or `.pfx` file, and `keystorePassword` must unlock
the keystore. `keyPassword` defaults to `keystorePassword`, `digestAlgorithm` defaults to
`SHA256`, and `alias` may be omitted to use the sole keystore entry or the first key entry POI can
resolve. `pkcs12Path` follows the same request-owned path rule as other request file paths.

When the CLI reads the request via `--request <path>`, relative persistence `path` values resolve
from that request file's directory. When the request JSON arrives on stdin, pass
`--execution-root <path>` and those same relative paths resolve from that explicit directory.

The save path must end in `.xlsx`.

The response uses one failure-capable save result:
- `requestedPath` — the literal `path` string from the request.
- `write.status=WRITTEN` plus `write.executionPath` when the file was actually written.
- `write.status=NOT_WRITTEN` when the run failed before any file write happened.

For every writing `EXISTING` source request, `persistence.security` must declare both encryption
and signature policies. There is no implicit source-encryption preservation or source-signature
carry-forward. Use `PRESERVE_SOURCE` only for a source that is encrypted with a supported AGILE
write envelope; use `SIGN` when the output must carry a package signature. A writing `NEW` source
may omit `security`; its declared default is `{ "encryption": { "type": "NONE" }, "signature":
{ "type": "NONE" } }`.

`ifExists=REJECT` requires the destination path to be absent. `ifExists=REPLACE` enables create-or-replace behavior while preserving the same `requestedPath` versus `executionPath` response split.

They are identical when an absolute path with no `..` segments is supplied. They differ when a
relative path (e.g. `"report.xlsx"`) or a path containing `..` segments is used.

Successful `SAVE_AS` responses therefore look like:

```json
{
  "type": "SAVE_AS",
  "requestedPath": "out/report.xlsx",
  "write": {
    "status": "WRITTEN",
    "executionPath": "/work/out/report.xlsx"
  }
}
```

If the run fails before persistence, the same `SAVE_AS` intent still appears on the failure
response, but the write result becomes:

```json
{
  "type": "SAVE_AS",
  "requestedPath": "out/report.xlsx",
  "write": {
    "status": "NOT_WRITTEN"
  }
}
```

```json
{
  "type": "OVERWRITE"
}
```
Overwrite the source file (requires `source.type=EXISTING`). `OVERWRITE` does not accept its own
`path` field; it always writes back to `source.path`. The response includes `sourcePath` (the
original source path string) whenever an `EXISTING` source path was available, and otherwise keeps
`type=OVERWRITE` while omitting `sourcePath`. In both cases it carries the same failure-capable
`write` result described above.

```json
{
  "type": "OVERWRITE",
  "security": {
    "encryption": {
      "type": "PRESERVE_SOURCE"
    },
    "signature": {
      "type": "SIGN",
      "signature": {
        "pkcs12Path": "signing-material.p12",
        "keystorePassword": "changeit",
        "alias": "gridgrind-signing"
      }
    }
  }
}
```

Every `OVERWRITE` request against an existing source declares the final encryption and signature
state. Use `signature.type=SIGN` to replace a source signature with a new signature, or
`signature.type=NONE` to make an intentional unsigned output.

Use `{ "type": "NONE" }` to run mutations, assertions, and inspections without saving.

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
      "font": { "fontName": "Aptos", "fontColor": { "type": "RGB", "rgb": "#44546A" } }
    },
    {
      "source": { "type": "INLINE", "text": "Budget" },
      "font": { "bold": true, "fontColor": { "type": "RGB", "rgb": "#C00000" } }
    }
  ]
}
{ "type": "NUMBER",    "number": 8.40                }
{ "type": "BOOLEAN",   "bool": true                  }
{ "type": "FORMULA",   "source": { "type": "INLINE", "text": "SUM(B2:B3)" } }
{ "type": "RAW_FORMULA", "source": { "type": "INLINE", "text": "LAMBDA(x,x+1)(A1)" } }
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
`FORMULA` and `RAW_FORMULA` text is the OOXML `<f>` body: it must not be empty or begin with `=`.
`FORMULA` payloads are scalar only. Array-formula braces such as `{=SUM(A1:A2*B1:B2)}` are
rejected as `INVALID_FORMULA`. `RAW_FORMULA` persists opaque XML-safe formula character data
without routing the body through POI's write parser, so it is the explicit route for newer Excel
syntax such as `LAMBDA` and `LET`. Invalid opaque framing or XML 1.0-forbidden character data is
rejected as `INVALID_FORMULA_TEXT`. Loaded formulas that POI parses but cannot evaluate are kept
unchanged under lenient evaluation and surface `FORMULA_NOT_EVALUATED`; strict evaluation fails.
Normal `FORMULA` text is validated when its mutation executes, after all preceding mutations are
present. If a later calculation, inspection, or workbook operation surfaces that formula's failure,
the primary `problem.context.step` is the authoring mutation and optional
`problem.context.surfacedAtStep` identifies the later trigger. `RAW_FORMULA` remains opaque in
this path and never receives evaluator-driven attribution.
Factual formula reads do not invoke the evaluator: use `GET_FORMULA_SURFACE` for grouped authored
bodies or a `GET_CELLS` projection containing only `FORMULA` for exact cell text. Adding `VALUE`
or `FORMAT` deliberately requests evaluator-backed output.

---
