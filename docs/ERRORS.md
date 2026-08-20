---
afad: "4.0"
version: "0.72.0"
domain: ERRORS
updated: "2026-08-08"
route:
  keywords: [gridgrind, errors, problem, code, category, recovery, failure, assertion-failed, invalid-json, invalid-request-shape, invalid-formula, unsupported-formula-construct, sheet-not-found, named-range-not-found, workbook-not-found, workbook-password-required, invalid-workbook-password, invalid-signing-configuration, workbook-security-error, input-source-not-found, input-source-unavailable, input-source-io-error, source-backed, standard_input, utf8_file, file, causes, context, sourceType, persistenceType, coordinates, rowindex, columnindex]
  questions: ["what error codes does gridgrind return", "what does a gridgrind failure response look like", "how do I handle gridgrind errors", "what is the problem model", "how do I read gridgrind error context", "how do I interpret gridgrind row or column index errors", "how does gridgrind report assertion failures", "how does gridgrind report encrypted workbook password failures", "how does gridgrind report signing failures", "how does gridgrind report source-backed input failures", "what happens if a gridgrind input file is missing"]
---

# Error Reference

**Purpose**: Problem codes, categories, and the full error response model.
**Prerequisites**: [README](../README.md) for the basic response structure.

---

## Failure Response Shape

```json
{
  "protocolVersion": "V2",
  "status": "FAILED",
  "planId": "set-total-pass",
  "persistence": {
    "type": "SAVE_AS",
    "requestedPath": "out/budget-reviewed.xlsx",
    "write": {
      "status": "NOT_WRITTEN"
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
      "finishedAt": "2026-04-19T09:30:08Z",
      "durationMillis": 8
    },
    "calculation": {
      "preflight": {
        "status": "NOT_REQUESTED",
        "durationMillis": 0
      },
      "execution": {
        "status": "NOT_REQUESTED",
        "durationMillis": 0
      }
    },
    "persistencePhase": {
      "status": "NOT_STARTED",
      "durationMillis": 0
    },
    "close": {
      "status": "SUCCEEDED",
      "startedAt": "2026-04-19T09:30:09Z",
      "finishedAt": "2026-04-19T09:30:10Z",
      "durationMillis": 1
    },
    "steps": [
      {
        "stepIndex": 2,
        "stepId": "set-total",
        "stepKind": "MUTATION",
        "stepType": "SET_CELL",
        "resolvedTargets": [
          {
            "kind": "CELL",
            "label": "Cell Inventory!ZZZ999999"
          }
        ],
        "phase": {
          "status": "FAILED",
          "startedAt": "2026-04-19T09:30:08Z",
          "finishedAt": "2026-04-19T09:30:09Z",
          "durationMillis": 1
        },
        "outcome": "FAILED",
        "failure": {
          "code": "INVALID_CELL_ADDRESS",
          "category": "REQUEST",
          "stage": "EXECUTE_STEP",
          "message": "Cell address 'ZZZ999999' is not a valid A1-notation address."
        }
      }
    ],
    "outcome": {
      "status": "FAILED",
      "plannedStepCount": 4,
      "completedStepCount": 2,
      "durationMillis": 10,
      "failedStepIndex": 2,
      "failedStepId": "set-total",
      "failureCode": "INVALID_CELL_ADDRESS"
    }
  },
  "calculation": {
    "policy": {
      "strategy": {
        "type": "DO_NOT_CALCULATE"
      }
    },
    "execution": {
      "status": "NOT_REQUESTED",
      "evaluatedFormulaCount": 0,
      "cachesCleared": false,
      "markRecalculateOnOpenApplied": false
    }
  },
  "warnings": [],
  "assertions": [],
  "inspections": [],
  "problem": {
    "code": "INVALID_CELL_ADDRESS",
    "category": "REQUEST",
    "recovery": "CHANGE_REQUEST",
    "title": "Invalid cell address",
    "message": "Cell address 'ZZZ999999' is not a valid A1-notation address.",
    "resolution": "Use a valid A1-style address such as A1 or BC12.",
    "context": {
      "stage": "EXECUTE_STEP",
      "stepIndex": 2,
      "stepId": "set-total",
      "stepKind": "MUTATION",
      "stepType": "SET_CELL",
      "sheetName": "Inventory",
      "address": "ZZZ999999"
    },
    "causes": [
      {
        "code": "INVALID_CELL_ADDRESS",
        "message": "Cell address 'ZZZ999999' is not a valid A1-notation address.",
        "stage": "EXECUTE_STEP"
      }
    ]
  }
}
```

The top-level `persistence` block is always present on failed responses too. It preserves the
requested save mode and intended path even when the run failed before any write happened, in which
case `write.status=NOT_WRITTEN`. The `journal` block is always present as well. It records
top-level phase timing plus ordered per-step outcomes even when the request fails before
persistence. Source-backed text and binary loading runs first under `journal.inputResolution`,
before the workbook is opened.
Persist-workbook failures point at `sourceWorkbookPath` or `persistencePath` inside
`problem.context`; they do not repeat a nested `problem.context.persistence` object.

Pre-execution command failures — CLI argument errors, help-routing, discovery-lookup failures, and
request read/parse/validate errors (including static semantic `INVALID_REQUEST` failures, as well
as `INVALID_ENCODING`, `INVALID_JSON`, and `INVALID_REQUEST_SHAPE` intake failures caught before
workbook execution begins) —
use the plural `CommandError` envelope instead of an execution result:

```json
{
  "protocolVersion": "V2",
  "command": "execute",
  "status": "REJECTED",
  "problems": [
    {
      "code": "INVALID_ARGUMENTS",
      "category": "ARGUMENTS",
      "recovery": "CHANGE_REQUEST",
      "title": "Invalid CLI arguments",
      "message": "Unknown argument: --bogus",
      "resolution": "Use one exact CLI flag. Start from --help for the synopsis, --help-protocol for the grammar, or --help-guidance for workflow-oriented commands.",
      "context": {
        "stage": "PARSE_ARGUMENTS",
        "argument": {
          "type": "NAMED",
          "argument": "--bogus"
        }
      },
      "causes": []
    }
  ]
}
```

`problems` is always nonempty. Argument and operational CLI failures naturally contain one entry;
request intake can carry every independently observable structural or constructor-level binding
finding. `REJECTED` means workbook execution never began; it does not describe the problem's fault
domain. Read `problems[*].category` for that domain.

`WorkbookResult` is emitted only by the execution command after workbook execution begins. Its
`status` is `SUCCEEDED` or `FAILED`; `problem` appears only on `FAILED`. `--doctor-request` keeps
its own successful `RequestDoctorReport` payload: findings make that report `valid:false`, not a
`CommandError`. For the same request bytes, execution rejection and doctor findings use the same
ordered problem core.

Every emitted `problems[]` and `warnings[]` uses one deterministic order: pipeline phase, present
UTF-8 byte offset before token-less locations, step index, duplicate occurrence, code, then an
internal phase-local allocation tie-breaker. That last tie-breaker is never serialized. CLI argument
details remain in `problems[*].context.argument`; request cursor details remain in
`problems[*].context.json`; category, title, resolution, and causes remain in `problems[*]`.

Every entry in `problem.causes` also carries an explicit `stage` token. Cause diagnostics are not
stage-less fallbacks; they preserve the same pipeline stage vocabulary used by the primary
`problem.context.stage` classification.

Assertion mismatches attach an additional `problem.assertionFailure` payload:

```json
{
  "status": "FAILED",
  "protocolVersion": "V2",
  "problem": {
    "code": "ASSERTION_FAILED",
    "category": "ASSERTION",
    "recovery": "CHANGE_REQUEST",
    "title": "Assertion failed",
    "message": "EXPECT_CELL_VALUE mismatched effective values at B2",
    "resolution": "Inspect problem.assertionFailure observations, then adjust the failing assertion or preceding workbook mutations and retry.",
    "context": {
      "stage": "EXECUTE_STEP",
      "stepIndex": 3,
      "stepId": "assert-total",
      "stepType": "EXPECT_CELL_VALUE",
      "sheetName": "Budget",
      "address": "B2"
    },
    "assertionFailure": {
      "stepId": "assert-total",
      "assertionType": "EXPECT_CELL_VALUE",
      "target": {
        "type": "CELL_BY_ADDRESS",
        "sheetName": "Budget",
        "address": "B2"
      },
      "assertion": {
        "type": "EXPECT_CELL_VALUE",
        "expectedValue": {
          "type": "NUMBER",
          "number": 1200
        }
      },
      "observations": [
        {
          "type": "GET_CELLS",
          "stepId": "assert-total",
          "sheetName": "Budget",
          "cells": [
            {
              "type": "NUMBER",
              "address": "B2",
              "numberValue": 900.0
            }
          ]
        }
      ]
    },
    "causes": []
  }
}
```

---

## Problem Codes

### Arguments (`ARGUMENTS` category)

| Code | Trigger |
|:-----|:--------|
| `INVALID_ARGUMENTS` | Unrecognized or malformed CLI arguments (e.g. unknown flag, missing value). |

### Request (`REQUEST` category)

| Code | Trigger |
|:-----|:--------|
| `INVALID_ENCODING` | Request bytes are not valid UTF-8, so GridGrind cannot begin JSON parsing. |
| `INVALID_JSON` | Request payload is not syntactically valid JSON. |
| `INVALID_REQUEST_SHAPE` | JSON is syntactically valid, but fields, discriminator IDs, explicit `null` placeholders, or token shapes do not match the GridGrind protocol schema. Messages are product-owned, classify the failure structurally at intake from the effective creator/discriminator contract, and point `context.jsonPath` at the exact offending field without leaking Jackson or Java class names. |
| `INPUT_SOURCE_UNAVAILABLE` | A source-backed authored field requested `STANDARD_INPUT`, but no stdin bytes were bound for authored input content. On the CLI this usually means the request itself was also read from stdin instead of `--request <path>`. |
| `INVALID_REQUEST` | JSON is valid and binds successfully, but the parsed request violates GridGrind business or cross-field validation, including non-`.xlsx` workbook paths, invalid `MOVE_SHEET` indexes, invalid/conflicting `RENAME_SHEET` targets, invalid hyperlink/comment/named-range payloads, invalid structural layout values, signed-workbook persistence requests that mutate the workbook without explicit `persistence.security.signature`, encrypted-source persistence that would implicitly carry forward a non-authorable OOXML write envelope, or `UNMERGE_CELLS` requests that do not match an existing merged region exactly. These request-owned invariants preserve exact offending-field paths and cause-specific resolutions on the public problem surface. |
| `INVALID_CELL_ADDRESS` | A1-notation cell address is malformed. |
| `INVALID_RANGE_ADDRESS` | A1-notation range is malformed or its dimensions do not match `rows`, including invalid `MERGE_CELLS` or `UNMERGE_CELLS` ranges. |

### Assertion (`ASSERTION` category)

| Code | Trigger |
|:-----|:--------|
| `ASSERTION_FAILED` | One authored assertion step did not match the observed workbook state. The failure includes `problem.assertionFailure` with the failed assertion contract and the observed factual read payloads that caused the mismatch. Entity-presence assertions (`EXPECT_SHEET_PRESENT`, `EXPECT_SHEET_ABSENT`, `EXPECT_NAMED_RANGE_PRESENT`, `EXPECT_NAMED_RANGE_ABSENT`, `EXPECT_TABLE_PRESENT`, `EXPECT_TABLE_ABSENT`, `EXPECT_PIVOT_TABLE_PRESENT`, `EXPECT_PIVOT_TABLE_ABSENT`, `EXPECT_CHART_PRESENT`, `EXPECT_CHART_ABSENT`) treat selector misses as zero observed entities instead of surfacing selector-specific `*_NOT_FOUND` errors. |

### Formula (`FORMULA` category)

| Code | Trigger |
|:-----|:--------|
| `INVALID_FORMULA` | Formula syntax is not valid Excel formula syntax on the current request path. Scalar `SET_CELL` / `SET_RANGE` `FORMULA` values reject request-authored array-formula braces such as `{=...}`; use `SET_ARRAY_FORMULA` for contiguous array groups. |
| `UNSUPPORTED_FORMULA_CONSTRUCT` | The authored formula uses a valid Excel construct that Apache POI cannot parse on the write path. Authored `LAMBDA` and `LET` currently surface here. |
| `MISSING_EXTERNAL_WORKBOOK` | Formula evaluation needs an external workbook binding that was not supplied and cached-value fallback is not enabled. |
| `UNREGISTERED_USER_DEFINED_FUNCTION` | Formula evaluation encountered a UDF that is not registered in `formulaEnvironment`. |
| `UNSUPPORTED_FORMULA` | Formula syntax is valid and Apache POI can load it, but the function or construct is not supported by Apache POI's evaluator. |

### Resource (`RESOURCE` category)

| Code | Trigger |
|:-----|:--------|
| `WORKBOOK_NOT_FOUND` | `source.type=EXISTING` path does not exist. |
| `INPUT_SOURCE_NOT_FOUND` | A source-backed authored field referenced a `UTF8_FILE` or `FILE` path that does not exist. When the CLI read the request via `--request <path>`, relative authored-input paths were resolved from that request file's directory. When the request JSON arrived on stdin, the same authored-input paths were resolved from the explicit `--execution-root <path>`, not from ambient process state. |
| `SHEET_NOT_FOUND` | A step target or nested payload references a sheet that does not exist. This can surface across sheet-backed writes and reads, layout or structure edits against existing sheets, table or pivot definitions, drawing selectors, and formula-evaluation targets. Use `ENSURE_SHEET` only for create-before-write flows; it does not replace references to already existing sheet names elsewhere in the request. |
| `NAMED_RANGE_NOT_FOUND` | A named-range inspection selector or delete step references a workbook- or sheet-scoped name that does not exist. |
| `CELL_NOT_FOUND` | The request named a cell that does not physically exist for a workflow that requires a real stored cell. The current public path is `execution.calculation.strategy=EVALUATE_TARGETS`: every addressed target must point at an existing formula cell. By contrast, `GET_CELLS` returns blank snapshots for unwritten cells, and `CLEAR_HYPERLINK` / `CLEAR_COMMENT` stay no-ops when the cell does not physically exist. |

### Security (`SECURITY` category)

| Code | Trigger |
|:-----|:--------|
| `WORKBOOK_PASSWORD_REQUIRED` | `source.type=EXISTING` points to an encrypted OOXML workbook and `source.security.password` was omitted. |
| `INVALID_WORKBOOK_PASSWORD` | `source.security.password` was supplied for an encrypted OOXML workbook, but it did not decrypt the package. |
| `INVALID_SIGNING_CONFIGURATION` | `persistence.security.signature` did not point to a readable PKCS#12 keystore or the configured alias/password/digest settings could not be resolved. |
| `WORKBOOK_SECURITY_ERROR` | OOXML cryptographic inspection, encryption, or signing failed after request validation due to package or runtime security state. |

### I/O (`IO` category)

| Code | Trigger |
|:-----|:--------|
| `INPUT_SOURCE_IO_ERROR` | A source-backed authored field pointed at a file that exists but could not be read, or stdin-backed source bytes could not be consumed cleanly. |
| `IO_ERROR` | File could not be read or written. Resolutions are stage-specific: `OPEN_WORKBOOK` points at the source workbook path, `PERSIST_WORKBOOK` distinguishes overwrite versus `SAVE_AS` destinations and calls out `SAVE_AS.ifExists=REJECT` collisions separately from broader write failures, and `WRITE_RESPONSE` points at the authored `--response` path. Transport-owned write failures preserve the attempted path and, when available, the operating-system reason. |

### Internal (`INTERNAL` category)

| Code | Trigger |
|:-----|:--------|
| `INTERNAL_ERROR` | Unexpected engine failure not covered by the above codes. |

---

## Categories

| Category | Meaning |
|:---------|:--------|
| `ARGUMENTS` | CLI argument was unrecognized or malformed. Fix the command invocation. |
| `REQUEST` | Request JSON is malformed, does not match the protocol shape, or violates semantic validation. Fix the request. |
| `ASSERTION` | An authored verification step did not match the observed workbook state. Inspect `problem.assertionFailure.observations`, then fix the expectation or the authored mutations. |
| `FORMULA` | Formula syntax is invalid, evaluation is missing required external/UDF configuration, or the construct is outside Apache POI's parser/evaluator support. Fix the formula or evaluator setup. |
| `RESOURCE` | Referenced workbook, sheet, or cell does not exist. Fix the path or name. |
| `SECURITY` | Workbook encryption, password, or OOXML signing failed. Fix the password or signing configuration, or inspect the workbook package and runtime crypto environment. |
| `IO` | Filesystem failure reading or writing a file. Check paths, permissions, and disk state. |
| `INTERNAL` | Unexpected engine error. Capture details and escalate. |

---

## Recovery Strategies

| Recovery | Suggested Action |
|:---------|:----------------|
| `CHANGE_REQUEST` | Fix the failing field or argument and resubmit. |
| `CHECK_ENVIRONMENT` | Verify file paths, permissions, disk space, and file locks before retrying. For `WORKBOOK_SECURITY_ERROR`, also inspect the workbook package and the runtime cryptographic environment. |
| `ESCALATE` | Internal error — capture the full problem object and escalate. |

---

## Context Fields

The `context` block provides structured metadata about where the failure occurred:

| Field | Description |
|:------|:------------|
| `stage` | `PARSE_ARGUMENTS`, `CLI_RUNTIME`, `READ_REQUEST`, `BIND_REQUEST`, `VALIDATE_REQUEST`, `RESOLVE_INPUTS`, `OPEN_WORKBOOK`, `EXECUTE_STEP`, `CALCULATION_PREFLIGHT`, `CALCULATION_EXECUTION`, `PERSIST_WORKBOOK`, `EXECUTE_REQUEST`, `WRITE_RESPONSE` |
| `argument` | The CLI flag or argument token that failed parsing, when the stage is `PARSE_ARGUMENTS`. |
| `requestPath` | The request file path used for `READ_REQUEST` or `BIND_REQUEST`, when the CLI read JSON from `--request <path>`. |
| `sourceType` | Request `source.type` when the failure occurred after request parsing, including `EXECUTE_REQUEST` failures. |
| `persistenceType` | Request `persistence.type` when the failure occurred after request parsing, including `EXECUTE_REQUEST` failures. |
| `sourceWorkbookPath` | The workbook path involved in `OPEN_WORKBOOK` or `PERSIST_WORKBOOK`, when a source workbook path exists. |
| `persistencePath` | The persistence destination path involved in `PERSIST_WORKBOOK`, when one exists. |
| `inputKind` | Authored source-backed field family when the failure occurred during `RESOLVE_INPUTS`, for example `cell text`, `picture payload`, or `embedded object preview image`. |
| `inputPath` | Authored `UTF8_FILE` or `FILE` path when the failure occurred during `RESOLVE_INPUTS`, if the failing source referenced a path. |
| `stepIndex` | Zero-based index of the failing step in `steps`. |
| `stepId` | Caller-defined step correlation ID for the failing step. |
| `stepKind` | High-level step family of the failing step: `MUTATION`, `ASSERTION`, or `INSPECTION`. |
| `stepType` | The action, assertion, or query `type` field of the failing step (for example `SET_CELL`, `EXPECT_CELL_VALUE`, or `GET_CELLS`). |
| `sheetName` | Sheet referenced by the failing step, if applicable. |
| `address` | Cell address, if applicable. |
| `range` | Range, if applicable. |
| `formula` | Formula text, if applicable. |
| `namedRangeName` | Named range involved in the failure, if applicable. |
| `jsonPath` | GridGrind dotted JSON path to the offending request value itself, such as `steps[0].action.zoomPercent`, `steps[0].target.type`, `steps[0].action.rows.cells[0][1].source.path`, or `protocolVersion`. It is present for request intake, static validation, and source-backed input-resolution failures when an authored request location is available. |
| `jsonLine` | One-based line number in the request payload when request parsing exposed a concrete cursor. |
| `jsonColumn` | One-based column number in the request payload when request parsing exposed a concrete cursor. |
| `responsePath` | The response file path that failed during `WRITE_RESPONSE`, when the CLI was writing to `--response <path>`. |

For the exact machine shape, four context stages carry typed nested helpers rather than flattened
nullable fields:

- `PARSE_ARGUMENTS` uses `context.argument`, where `type=UNKNOWN|NAMED` and the concrete flag or
  operand lives at `context.argument.argument` when `type=NAMED`.
- `CLI_RUNTIME` identifies a last-resort CLI failure before workbook execution when no narrower
  command, request, or workbook stage truthfully owns the fault.
- `READ_REQUEST` uses `context.request` (`STANDARD_INPUT` or `FILE`) plus `context.json`
  (`UNAVAILABLE`, `PATH_ONLY`, `BYTE_OFFSET`, `PATH_BYTE_OFFSET`, `DUPLICATE_KEY`,
  `LINE_COLUMN`, or `LOCATED`). `PATH_BYTE_OFFSET` preserves both the owned path and its exact
  zero-based UTF-8 byte offset. Object-member failures, including scalar-shape and explicit-null
  findings, point to the property's opening quote; array-element and root-value failures point to
  the value token. `DUPLICATE_KEY` preserves the
  containing-object path, key, duplicate occurrence ordinal, and property-token byte offset
  without inventing a false JSON path. Every authored occurrence is structurally checked, so a
  repeated known field produces its duplicate-key finding and any independently provable problem
  with that repeated value, such as explicit `null` or the wrong scalar kind.
- `BIND_REQUEST` uses the same `context.request` and `context.json` shape for a syntactically and
  structurally valid JSON fragment that cannot satisfy its request-model creator contract. This
  keeps structural intake defects before constructor-level binding defects in deterministic
  diagnostic order.
- `WRITE_RESPONSE` uses `context.output` (`STANDARD_OUTPUT` or `FILE`).

Without `--response`, the primary `CommandError`, `WorkbookResult`, doctor report, or discovery
payload is the sole stdout content. With `--response <path>`, that primary payload is written only
to the requested file. GridGrind does not mirror an equivalent diagnostic on stderr. If that file
cannot be written and stdout is writable, the already-rendered primary payload goes to stdout unchanged and stderr
receives exactly one transport-only JSON line, `{"wroteTo":"STDOUT","responsePath":"..."}`.
The notice has no status, exit code, or problem data and is not a second result schema. If stdout
is unavailable or fails while a payload is being written, GridGrind exits nonzero without moving
or retrying the primary payload on another channel.

Values declared `secret: true` in the request contract are protected by their exact JSON owner
path. A binding or validation problem at a declared secret path uses a generic sensitive-safe
message, and a structured problem carrying that same path redacts its message, resolution, and
causes. GridGrind deliberately does not use global string replacement: a short password must not
alter unrelated workbook data or diagnostics that merely contain the same text. Live journal events
do not carry request-field values and remain the authored engine events. Last-resort failures that
classify as `INTERNAL_ERROR` use the canonical internal-error title rather than reproducing
arbitrary throwable text.

## Index-Based Validation Messages

Row and column validation failures report both the raw zero-based value and the Excel-native
equivalent. For example:

- `firstRowIndex 5 (Excel row 6)`
- `firstColumnIndex 5 (Excel column F)`

This applies to structural edit bounds, print-title band validation, and related index-based
operations. `address` and `range` fields still use plain A1 notation.

---

## Diagnostic Causes

`causes` is an ordered list of GridGrind-classified diagnostic entries. The first entry describes
the primary classified failure in GridGrind terms, and later entries capture supplemental failures
that occurred in other stages while GridGrind was already handling the main problem.

Each entry carries:

| Field | Description |
|:------|:------------|
| `code` | Stable GridGrind problem code for this diagnostic entry. |
| `message` | Product-owned diagnostic message for this entry. |
| `stage` | Pipeline stage where this diagnostic originated; omitted when the entry is not attributed to a stage. |

Agents should inspect `code` to distinguish between, for example, `INVALID_FORMULA`,
`UNSUPPORTED_FORMULA_CONSTRUCT`, `MISSING_EXTERNAL_WORKBOOK`,
`UNREGISTERED_USER_DEFINED_FUNCTION`, and `UNSUPPORTED_FORMULA` without depending on Java
exception class names or parser-library details.

## Assertion Failure Payload

When `problem.code=ASSERTION_FAILED`, `problem.assertionFailure` is always present. It carries:

| Field | Description |
|:------|:------------|
| `stepId` | The authored assertion step ID that failed. |
| `assertionType` | The stable assertion discriminator such as `EXPECT_CELL_VALUE` or `ALL_OF`. |
| `target` | The selector payload the assertion executed against. |
| `assertion` | The authored assertion contract itself. |
| `observations` | Ordered factual inspection results gathered by GridGrind while evaluating the assertion. These are the authoritative mismatch facts to inspect before retrying. |
