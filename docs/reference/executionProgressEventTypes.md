<!-- Generated from core metadata; do not edit. -->
# executionProgressEventTypes

## STARTED

Compact JSONL lifecycle event emitted when one execution phase starts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| timestamp | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| category | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| stepIndex | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stepId | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "STARTED",
  "summary" : "Compact JSONL lifecycle event emitted when one execution phase starts.",
  "fields" : [ {
    "name" : "timestamp",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "category",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "PLAN", "VALIDATION", "RESOLVE_INPUTS", "OPEN", "CALCULATION_PREFLIGHT", "CALCULATION_EXECUTION", "STEP", "PERSIST", "CLOSE" ]
  }, {
    "name" : "stepIndex",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stepId",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## SUCCEEDED

Compact JSONL lifecycle event emitted when one execution phase succeeds.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| timestamp | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| category | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| stepIndex | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stepId | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "SUCCEEDED",
  "summary" : "Compact JSONL lifecycle event emitted when one execution phase succeeds.",
  "fields" : [ {
    "name" : "timestamp",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "category",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "PLAN", "VALIDATION", "RESOLVE_INPUTS", "OPEN", "CALCULATION_PREFLIGHT", "CALCULATION_EXECUTION", "STEP", "PERSIST", "CLOSE" ]
  }, {
    "name" : "stepIndex",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stepId",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## FAILED

Compact JSONL lifecycle event emitted when one execution phase fails with one classified problem code.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| timestamp | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| category | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| problemCode | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| stepIndex | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stepId | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FAILED",
  "summary" : "Compact JSONL lifecycle event emitted when one execution phase fails with one classified problem code.",
  "fields" : [ {
    "name" : "timestamp",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "category",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "PLAN", "VALIDATION", "RESOLVE_INPUTS", "OPEN", "CALCULATION_PREFLIGHT", "CALCULATION_EXECUTION", "STEP", "PERSIST", "CLOSE" ]
  }, {
    "name" : "problemCode",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "INVALID_ARGUMENTS", "INVALID_JSON", "INVALID_ENCODING", "INVALID_REQUEST_SHAPE", "INVALID_REQUEST", "NUMBER_NOT_REPRESENTABLE", "PATH_ESCAPES_ROOT", "ASSERTION_FAILED", "PRESERVATION_FAILED", "INVALID_CELL_ADDRESS", "INVALID_RANGE_ADDRESS", "INVALID_FORMULA", "INVALID_FORMULA_TEXT", "UNSUPPORTED_FORMULA_CONSTRUCT", "CIRCULAR_FORMULA_REFERENCE", "MISSING_EXTERNAL_WORKBOOK", "UNREGISTERED_USER_DEFINED_FUNCTION", "UNSUPPORTED_FORMULA", "WORKBOOK_NOT_FOUND", "INPUT_SOURCE_NOT_FOUND", "INPUT_SOURCE_UNAVAILABLE", "INPUT_SOURCE_IO_ERROR", "SHEET_NOT_FOUND", "NAMED_RANGE_NOT_FOUND", "CELL_NOT_FOUND", "WORKBOOK_NOT_OPENABLE", "SOURCE_PATH_IS_DIRECTORY", "OUTPUT_PATH_IS_DIRECTORY", "OUTPUT_PATH_ALREADY_EXISTS", "WORKBOOK_PASSWORD_REQUIRED", "INVALID_WORKBOOK_PASSWORD", "ENCRYPTION_SOURCE_NOT_ENCRYPTED", "ENCRYPTION_SOURCE_NOT_PRESERVABLE", "INVALID_SIGNING_CONFIGURATION", "WORKBOOK_SECURITY_ERROR", "UNSAFE_PATH_ACCESS", "AUTHORITY_DENIED", "PUBLICATION_UNCERTAIN", "IO_ERROR", "INTERNAL_ERROR" ]
  }, {
    "name" : "stepIndex",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stepId",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
