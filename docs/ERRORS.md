---
afad: "3.3"
version: "1.0.0"
domain: ERRORS
updated: "2026-03-24"
route:
  keywords: [gridgrind, errors, problem, code, category, recovery, failure, invalid-json, invalid-formula, sheet-not-found, workbook-not-found, causes, context]
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
        "type": "InvalidCellAddressException",
        "className": "dev.erst.gridgrind.excel.InvalidCellAddressException",
        "message": "...",
        "stage": null
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
| `INVALID_JSON` | Request payload is not valid JSON. |
| `INVALID_REQUEST` | JSON is valid but the payload shape violates the protocol schema. |
| `INVALID_CELL_ADDRESS` | A1-notation cell address is malformed. |
| `INVALID_RANGE_ADDRESS` | A1-notation range is malformed or its dimensions do not match `rows`. |

### Formula (`FORMULA` category)

| Code | Trigger |
|:-----|:--------|
| `INVALID_FORMULA` | Formula syntax is not valid Excel formula syntax. |
| `UNSUPPORTED_FORMULA` | Formula syntax is valid but the function or construct is not supported by Apache POI. |

### Resource (`RESOURCE` category)

| Code | Trigger |
|:-----|:--------|
| `WORKBOOK_NOT_FOUND` | `source.mode=EXISTING` path does not exist. |
| `SHEET_NOT_FOUND` | An operation or analysis references a sheet that does not exist. |
| `CELL_NOT_FOUND` | An analysis `cells` entry references a cell that has not been written. |

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
| `REQUEST` | Request JSON is invalid or violates the protocol schema. Fix the request. |
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
| `stage` | `READ_REQUEST`, `VALIDATE_REQUEST`, `APPLY_OPERATION`, `ANALYZE_WORKBOOK`, `PERSIST_WORKBOOK`, `WRITE_RESPONSE` |
| `operationIndex` | Zero-based index of the failing operation in `operations`. |
| `operationType` | The `type` field of the failing operation (e.g. `SET_CELL`). |
| `sheetName` | Sheet referenced by the failing operation, if applicable. |
| `address` | Cell address, if applicable. |
| `range` | Range, if applicable. |
| `formula` | Formula text, if applicable. |
| `jsonPath` | JSON Pointer to the field that failed parsing (transport errors only). |
| `jsonLine` | Line number in the request payload (transport errors only). |
| `jsonColumn` | Column number in the request payload (transport errors only). |

---

## Causes Chain

`causes` is an ordered list of underlying exceptions from immediate cause to root cause.
Each entry carries:

| Field | Description |
|:------|:------------|
| `type` | Simple class name of the exception (e.g. `InvalidCellAddressException`). |
| `className` | Fully-qualified class name for precise matching. |
| `message` | Exception message. |
| `stage` | Pipeline stage where this cause originated, or `null` if not attributed to a stage. |

Agents can inspect `type` to distinguish between, for example, `FormulaException` (invalid
formula syntax) and `UnsupportedOperationException` (valid formula syntax, unsupported by POI).
