---
afad: "3.4"
version: "0.22.0"
domain: ERRORS
updated: "2026-03-31"
route:
  keywords: [gridgrind, errors, problem, code, category, recovery, failure, invalid-json, invalid-request-shape, invalid-formula, sheet-not-found, named-range-not-found, workbook-not-found, causes, context]
  questions: ["what error codes does gridgrind return", "what does a gridgrind failure response look like", "how do I handle gridgrind errors", "what is the problem model", "how do I read gridgrind error context"]
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
  "problem": {
    "code": "INVALID_CELL_ADDRESS",
    "category": "REQUEST",
    "recovery": "CHANGE_REQUEST",
    "title": "Invalid cell address",
    "message": "Cell address 'ZZZ999999' is not a valid A1-notation address.",
    "resolution": "Use a valid A1-style address such as A1 or BC12.",
    "context": {
      "stage": "APPLY_OPERATION",
      "operationIndex": 2,
      "operationType": "SET_CELL",
      "sheetName": "Inventory",
      "address": "ZZZ999999"
    },
    "causes": [
      {
        "code": "INVALID_CELL_ADDRESS",
        "message": "...",
        "stage": "APPLY_OPERATION"
      }
    ]
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
| `INVALID_REQUEST` | JSON is valid and binds successfully, but the parsed request violates GridGrind business or cross-field validation, including non-`.xlsx` workbook paths, invalid `MOVE_SHEET` indexes, invalid/conflicting `RENAME_SHEET` targets, invalid hyperlink/comment/named-range payloads, invalid structural layout values, or `UNMERGE_CELLS` requests that do not match an existing merged region exactly. |
| `INVALID_CELL_ADDRESS` | A1-notation cell address is malformed. |
| `INVALID_RANGE_ADDRESS` | A1-notation range is malformed or its dimensions do not match `rows`, including invalid `MERGE_CELLS` or `UNMERGE_CELLS` ranges. |

### Formula (`FORMULA` category)

| Code | Trigger |
|:-----|:--------|
| `INVALID_FORMULA` | Formula syntax is not valid Excel formula syntax. |
| `UNSUPPORTED_FORMULA` | Formula syntax is valid but the function or construct is not supported by Apache POI. |

### Resource (`RESOURCE` category)

| Code | Trigger |
|:-----|:--------|
| `WORKBOOK_NOT_FOUND` | `source.type=EXISTING` path does not exist. |
| `SHEET_NOT_FOUND` | An operation or read references a sheet that does not exist. All mutation operations (`SET_CELL`, `SET_RANGE`, `APPLY_STYLE`, `SET_HYPERLINK`, `CLEAR_HYPERLINK`, `SET_COMMENT`, `CLEAR_COMMENT`, `APPEND_ROW`, `AUTO_SIZE_COLUMNS`) require the sheet to exist; use `ENSURE_SHEET` first. |
| `NAMED_RANGE_NOT_FOUND` | A named-range read selector or delete request references a workbook- or sheet-scoped name that does not exist. |
| `CELL_NOT_FOUND` | Reserved. No current operation raises this code; `CLEAR_HYPERLINK` and `CLEAR_COMMENT` are no-ops when the cell does not physically exist, and `GET_CELLS` returns blank snapshots for unwritten cells. |

### I/O (`IO` category)

| Code | Trigger |
|:-----|:--------|
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
| `FORMULA` | Formula is syntactically invalid or unsupported by Apache POI. Fix the formula. |
| `RESOURCE` | Referenced workbook, sheet, or cell does not exist. Fix the path or name. |
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
| `stage` | `PARSE_ARGUMENTS`, `READ_REQUEST`, `VALIDATE_REQUEST`, `OPEN_WORKBOOK`, `APPLY_OPERATION`, `EXECUTE_READ`, `PERSIST_WORKBOOK`, `EXECUTE_REQUEST`, `WRITE_RESPONSE` |
| `sourceType` | Request `source.type` when the failure occurred after request parsing. |
| `persistenceType` | Request `persistence.type` when the failure occurred after request parsing. |
| `operationIndex` | Zero-based index of the failing operation in `operations`. |
| `operationType` | The `type` field of the failing operation (e.g. `SET_CELL`). |
| `readIndex` | Zero-based index of the failing read in `reads`. |
| `readType` | The `type` field of the failing read (e.g. `GET_CELLS`). |
| `requestId` | Caller-defined read correlation ID when the failure occurred during `EXECUTE_READ`. |
| `sheetName` | Sheet referenced by the failing operation, if applicable. |
| `address` | Cell address, if applicable. |
| `range` | Range, if applicable. |
| `formula` | Formula text, if applicable. |
| `namedRangeName` | Named range involved in the failure, if applicable. |
| `jsonPath` | JSON Pointer to the field that failed parsing (transport errors only). |
| `jsonLine` | Line number in the request payload (transport errors only). |
| `jsonColumn` | Column number in the request payload (transport errors only). |

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

Agents should inspect `code` to distinguish between, for example, `INVALID_FORMULA` and
`UNSUPPORTED_FORMULA`, without depending on Java exception class names or parser-library details.
