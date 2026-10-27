<!-- Generated from core metadata; do not edit. -->
# outcomeTypes

## SUCCEEDED

Execution outcome summary for a successful run.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| plannedStepCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| completedStepCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| durationMillis | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "SUCCEEDED",
  "summary" : "Execution outcome summary for a successful run.",
  "fields" : [ {
    "name" : "plannedStepCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "completedStepCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "durationMillis",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## FAILED

Execution outcome summary for a failed run, including the canonical failure code and an optional failing-step reference.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| plannedStepCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| completedStepCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| durationMillis | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| problemCode | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| failedStep | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"executionJournalFailureStepType"} |  |

```json
{
  "id" : "FAILED",
  "summary" : "Execution outcome summary for a failed run, including the canonical failure code and an optional failing-step reference.",
  "fields" : [ {
    "name" : "plannedStepCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "completedStepCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "durationMillis",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "problemCode",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "INVALID_ARGUMENTS", "INVALID_JSON", "INVALID_ENCODING", "INVALID_REQUEST_SHAPE", "INVALID_REQUEST", "NUMBER_NOT_REPRESENTABLE", "PATH_ESCAPES_ROOT", "ASSERTION_FAILED", "PRESERVATION_FAILED", "INVALID_CELL_ADDRESS", "INVALID_RANGE_ADDRESS", "INVALID_FORMULA", "INVALID_FORMULA_TEXT", "UNSUPPORTED_FORMULA_CONSTRUCT", "CIRCULAR_FORMULA_REFERENCE", "MISSING_EXTERNAL_WORKBOOK", "UNREGISTERED_USER_DEFINED_FUNCTION", "UNSUPPORTED_FORMULA", "WORKBOOK_NOT_FOUND", "INPUT_SOURCE_NOT_FOUND", "INPUT_SOURCE_UNAVAILABLE", "INPUT_SOURCE_IO_ERROR", "SHEET_NOT_FOUND", "NAMED_RANGE_NOT_FOUND", "CELL_NOT_FOUND", "WORKBOOK_NOT_OPENABLE", "SOURCE_PATH_IS_DIRECTORY", "OUTPUT_PATH_IS_DIRECTORY", "OUTPUT_PATH_ALREADY_EXISTS", "WORKBOOK_PASSWORD_REQUIRED", "INVALID_WORKBOOK_PASSWORD", "ENCRYPTION_SOURCE_NOT_ENCRYPTED", "ENCRYPTION_SOURCE_NOT_PRESERVABLE", "INVALID_SIGNING_CONFIGURATION", "WORKBOOK_SECURITY_ERROR", "UNSAFE_PATH_ACCESS", "AUTHORITY_DENIED", "PUBLICATION_UNCERTAIN", "IO_ERROR", "INTERNAL_ERROR" ]
  }, {
    "name" : "failedStep",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "executionJournalFailureStepType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
