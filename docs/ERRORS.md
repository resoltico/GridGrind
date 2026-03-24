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
    "category": "VALIDATION",
    "recovery": "CORRECT_REQUEST",
    "title": "Invalid cell address",
    "message": "Cell address 'ZZZ999999' is not a valid A1-notation address.",
    "resolution": "Correct the address in the failing operation and resubmit.",
    "context": {
      "stage": "APPLY_OPERATION",
      "operationIndex": 2,
      "operationType": "SET_CELL",
      "sheetName": "Inventory",
      "address": "ZZZ999999"
    },
    "causes": [
      { "type": "InvalidCellAddressException", "message": "..." }
    ]
  }
}
```

---

## Problem Codes

### Transport / Parsing

| Code | Trigger |
|:-----|:--------|
| `INVALID_JSON` | Request payload is not valid JSON. |
| `INVALID_REQUEST` | JSON is valid but payload shape violates the protocol schema. |

### Workbook Resource

| Code | Trigger |
|:-----|:--------|
| `WORKBOOK_NOT_FOUND` | `source.mode=EXISTING` path does not exist. |
| `WORKBOOK_OPEN_ERROR` | Existing file exists but cannot be opened (corrupt, wrong format, locked). |
| `WORKBOOK_SAVE_ERROR` | Workbook could not be saved to the specified path. |

### Sheet Resource

| Code | Trigger |
|:-----|:--------|
| `SHEET_NOT_FOUND` | An operation or analysis references a sheet that does not exist. |

### Cell / Range Addressing

| Code | Trigger |
|:-----|:--------|
| `INVALID_CELL_ADDRESS` | A1-notation cell address is malformed. |
| `INVALID_RANGE_ADDRESS` | A1-notation range is malformed or dimensions do not match `rows`. |

### Formula

| Code | Trigger |
|:-----|:--------|
| `INVALID_FORMULA` | Formula syntax is not valid Excel formula syntax. |
| `UNSUPPORTED_FORMULA` | Formula syntax is valid but the function or construct is not supported by Apache POI. |

### Internal

| Code | Trigger |
|:-----|:--------|
| `INTERNAL_ERROR` | Unexpected engine failure not covered by the above codes. |

---

## Categories

| Category | Meaning |
|:---------|:--------|
| `TRANSPORT` | Request could not be parsed. Fix the JSON. |
| `VALIDATION` | Request shape is invalid. Fix the request fields. |
| `RESOURCE` | Referenced workbook or sheet does not exist or cannot be accessed. |
| `EXECUTION` | Operation failed during workbook mutation. |
| `INTERNAL` | Unexpected engine error. |

---

## Recovery Strategies

| Recovery | Suggested Action |
|:---------|:----------------|
| `CORRECT_REQUEST` | Fix the failing field and resubmit. |
| `CHECK_FILE` | Verify the referenced file exists and is accessible. |
| `RETRY` | Transient resource failure — retry may succeed. |
| `CONTACT_SUPPORT` | Internal error — escalate. |

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

`causes` is an ordered list of underlying exceptions. Each entry carries:

| Field | Description |
|:------|:------------|
| `type` | Simple class name of the exception (e.g. `InvalidCellAddressException`). |
| `message` | Exception message. |

The chain starts with the immediate cause and ends with the root cause. Agents can inspect the
type chain to distinguish between, for example, `FormulaException` (invalid formula) and
`UnsupportedOperationException` (valid formula, unsupported by POI).
