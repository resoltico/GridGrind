<!-- Generated from core metadata; do not edit. -->
# formulaCapabilityReportType

## FormulaCapabilityReport

One classified formula-cell capability entry from calculation preflight.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cell | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"formulaCellTargetType"} |  |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| capability | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| problemCode | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| message | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FormulaCapabilityReport",
  "summary" : "One classified formula-cell capability entry from calculation preflight.",
  "fields" : [ {
    "name" : "cell",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "formulaCellTargetType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "capability",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "EVALUABLE_NOW", "UNEVALUABLE_NOW", "UNPARSEABLE_BY_POI" ]
  }, {
    "name" : "problemCode",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "INVALID_ARGUMENTS", "INVALID_JSON", "INVALID_ENCODING", "INVALID_REQUEST_SHAPE", "INVALID_REQUEST", "NUMBER_NOT_REPRESENTABLE", "PATH_ESCAPES_ROOT", "ASSERTION_FAILED", "PRESERVATION_FAILED", "INVALID_CELL_ADDRESS", "INVALID_RANGE_ADDRESS", "INVALID_FORMULA", "INVALID_FORMULA_TEXT", "UNSUPPORTED_FORMULA_CONSTRUCT", "CIRCULAR_FORMULA_REFERENCE", "MISSING_EXTERNAL_WORKBOOK", "UNREGISTERED_USER_DEFINED_FUNCTION", "UNSUPPORTED_FORMULA", "WORKBOOK_NOT_FOUND", "INPUT_SOURCE_NOT_FOUND", "INPUT_SOURCE_UNAVAILABLE", "INPUT_SOURCE_IO_ERROR", "SHEET_NOT_FOUND", "NAMED_RANGE_NOT_FOUND", "CELL_NOT_FOUND", "WORKBOOK_NOT_OPENABLE", "SOURCE_PATH_IS_DIRECTORY", "OUTPUT_PATH_IS_DIRECTORY", "OUTPUT_PATH_ALREADY_EXISTS", "WORKBOOK_PASSWORD_REQUIRED", "INVALID_WORKBOOK_PASSWORD", "ENCRYPTION_SOURCE_NOT_ENCRYPTED", "ENCRYPTION_SOURCE_NOT_PRESERVABLE", "INVALID_SIGNING_CONFIGURATION", "WORKBOOK_SECURITY_ERROR", "UNSAFE_PATH_ACCESS", "AUTHORITY_DENIED", "PUBLICATION_UNCERTAIN", "IO_ERROR", "INTERNAL_ERROR" ]
  }, {
    "name" : "message",
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
