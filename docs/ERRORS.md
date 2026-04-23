---
afad: "3.5"
version: "0.54.0"
domain: ERRORS
updated: "2026-04-22"
route:
  keywords: [gridgrind, errors, problem, code, category, recovery, failure, assertion-failed, invalid-json, invalid-request-shape, invalid-formula, sheet-not-found, named-range-not-found, workbook-not-found, workbook-password-required, invalid-workbook-password, invalid-signing-configuration, workbook-security-error, input-source-not-found, input-source-unavailable, input-source-io-error, source-backed, standard_input, utf8_file, file, causes, context, sourceType, persistenceType, coordinates, rowindex, columnindex]
  questions: ["what error codes does gridgrind return", "what does a gridgrind failure response look like", "how do I handle gridgrind errors", "what is the problem model", "how do I read gridgrind error context", "how do I interpret gridgrind row or column index errors", "how does gridgrind report assertion failures", "how does gridgrind report encrypted workbook password failures", "how does gridgrind report signing failures", "how does gridgrind report source-backed input failures", "what happens if a gridgrind input file is missing"]
---

# Error Reference

**Purpose**: Problem codes, categories, and the full error response model.
**Prerequisites**: [README](../README.md) for the basic response structure.

---

## Failure Response Shape

```json
{
  "status": "ERROR",
  "protocolVersion": "V1",
  "journal": {
    "planId": "set-total-pass",
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
      "finishedAt": "2026-04-19T09:30:08Z",
      "durationMillis": 8
    },
    "calculation": {
      "preflight": {
        "status": "NOT_STARTED",
        "durationMillis": 0
      },
      "execution": {
        "status": "NOT_STARTED",
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
    "warnings": [],
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

The `journal` block is always present. It records top-level phase timing plus ordered per-step
outcomes even when the request fails before persistence. Source-backed text and binary loading runs
first under `journal.inputResolution`, before the workbook is opened.

Assertion mismatches attach an additional `problem.assertionFailure` payload:

```json
{
  "status": "ERROR",
  "protocolVersion": "V1",
  "problem": {
    "code": "ASSERTION_FAILED",
    "category": "ASSERTION",
    "recovery": "CHANGE_REQUEST",
    "title": "Assertion failed",
    "message": "EXPECT_CELL_VALUE mismatched effective values at B2",
    "resolution": "Inspect the observed workbook facts, then adjust the plan expectations or authored mutations and retry.",
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
        "type": "BY_ADDRESS",
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
              "declaredType": "NUMBER",
              "value": 900,
              "displayValue": "900"
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
| `INVALID_JSON` | Request payload is not syntactically valid JSON. |
| `INVALID_REQUEST_SHAPE` | JSON is syntactically valid, but fields, discriminator IDs, or token shapes do not match the GridGrind protocol schema. Messages are product-owned and describe unknown fields, unknown type values, missing required fields, or wrong token shapes without leaking Jackson or Java class names. |
| `INPUT_SOURCE_UNAVAILABLE` | A source-backed authored field requested `STANDARD_INPUT`, but no stdin bytes were bound for authored input content. On the CLI this usually means the request itself was also read from stdin instead of `--request <path>`. |
| `INVALID_REQUEST` | JSON is valid and binds successfully, but the parsed request violates GridGrind business or cross-field validation, including non-`.xlsx` workbook paths, invalid `MOVE_SHEET` indexes, invalid/conflicting `RENAME_SHEET` targets, invalid hyperlink/comment/named-range payloads, invalid structural layout values, signed-workbook persistence requests that mutate the workbook without explicit `persistence.security.signature`, or `UNMERGE_CELLS` requests that do not match an existing merged region exactly. |
| `INVALID_CELL_ADDRESS` | A1-notation cell address is malformed. |
| `INVALID_RANGE_ADDRESS` | A1-notation range is malformed or its dimensions do not match `rows`, including invalid `MERGE_CELLS` or `UNMERGE_CELLS` ranges. |

### Assertion (`ASSERTION` category)

| Code | Trigger |
|:-----|:--------|
| `ASSERTION_FAILED` | One authored assertion step did not match the observed workbook state. The failure includes `problem.assertionFailure` with the failed assertion contract and the observed factual read payloads that caused the mismatch. Presence-style assertions (`EXPECT_PRESENT`, `EXPECT_ABSENT`) treat selector misses as zero observed entities instead of surfacing selector-specific `*_NOT_FOUND` errors. |

### Formula (`FORMULA` category)

| Code | Trigger |
|:-----|:--------|
| `INVALID_FORMULA` | Formula syntax is not valid Excel formula syntax on the current request path. Scalar `SET_CELL` / `SET_RANGE` `FORMULA` values reject request-authored array-formula braces such as `{=...}`; use `SET_ARRAY_FORMULA` for contiguous array groups. `LAMBDA` and `LET` are currently rejected as `INVALID_FORMULA` because Apache POI cannot parse them, and other newer constructs may fail the same way. |
| `MISSING_EXTERNAL_WORKBOOK` | Formula evaluation needs an external workbook binding that was not supplied and cached-value fallback is not enabled. |
| `UNREGISTERED_USER_DEFINED_FUNCTION` | Formula evaluation encountered a UDF that is not registered in `formulaEnvironment`. |
| `UNSUPPORTED_FORMULA` | Formula syntax is valid and Apache POI can load it, but the function or construct is not supported by Apache POI's evaluator. |

### Resource (`RESOURCE` category)

| Code | Trigger |
|:-----|:--------|
| `WORKBOOK_NOT_FOUND` | `source.type=EXISTING` path does not exist. |
| `INPUT_SOURCE_NOT_FOUND` | A source-backed authored field referenced a `UTF8_FILE` or `FILE` path that does not exist. |
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
| `IO_ERROR` | File could not be read or written — wrong path, missing permissions, disk full, or file locked. |

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
| `CHECK_ENVIRONMENT` | Verify file paths, permissions, disk space, and file locks before retrying. |
| `ESCALATE` | Internal error — capture the full problem object and escalate. |

---

## Context Fields

The `context` block provides structured metadata about where the failure occurred:

| Field | Description |
|:------|:------------|
| `stage` | `PARSE_ARGUMENTS`, `READ_REQUEST`, `VALIDATE_REQUEST`, `RESOLVE_INPUTS`, `OPEN_WORKBOOK`, `EXECUTE_STEP`, `CALCULATION_PREFLIGHT`, `CALCULATION_EXECUTION`, `PERSIST_WORKBOOK`, `EXECUTE_REQUEST`, `WRITE_RESPONSE` |
| `argument` | The CLI flag or argument token that failed parsing, when the stage is `PARSE_ARGUMENTS`. |
| `requestPath` | The request file path used for `READ_REQUEST`, when the CLI read JSON from `--request <path>`. |
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
| `jsonPath` | JSON Pointer to the field that failed parsing (transport errors only). |
| `jsonLine` | Line number in the request payload (transport errors only). |
| `jsonColumn` | Column number in the request payload (transport errors only). |
| `responsePath` | The response file path that failed during `WRITE_RESPONSE`, when the CLI was writing to `--response <path>`. |

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
| `stage` | Pipeline stage where this diagnostic originated, or `null` if not attributed to a stage. |

Agents should inspect `code` to distinguish between, for example, `INVALID_FORMULA`,
`MISSING_EXTERNAL_WORKBOOK`, `UNREGISTERED_USER_DEFINED_FUNCTION`, and `UNSUPPORTED_FORMULA`
without depending on Java exception class names or parser-library details.

## Assertion Failure Payload

When `problem.code=ASSERTION_FAILED`, `problem.assertionFailure` is always present. It carries:

| Field | Description |
|:------|:------------|
| `stepId` | The authored assertion step ID that failed. |
| `assertionType` | The stable assertion discriminator such as `EXPECT_CELL_VALUE` or `ALL_OF`. |
| `target` | The selector payload the assertion executed against. |
| `assertion` | The authored assertion contract itself. |
| `observations` | Ordered factual inspection results gathered by GridGrind while evaluating the assertion. These are the authoritative mismatch facts to inspect before retrying. |
