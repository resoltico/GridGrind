<!-- Generated from core metadata; do not edit. -->
# cellValueTypes

## BLANK

An effective blank cell value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "BLANK",
  "summary" : "An effective blank cell value.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## TEXT

One exact effective text value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TEXT",
  "summary" : "One exact effective text value.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## NUMBER

One effective numeric value. Nonfinite observed values retain their JSON string representation; authored writes and numeric expectations must be finite.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "NUMBER",
  "summary" : "One effective numeric value. Nonfinite observed values retain their JSON string representation; authored writes and numeric expectations must be finite.",
  "fields" : [ {
    "name" : "value",
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

## BOOLEAN

One exact effective boolean value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "BOOLEAN",
  "summary" : "One exact effective boolean value.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## ERROR

One GridGrind-owned effective error literal, including reported evaluation-only states.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ERROR",
  "summary" : "One GridGrind-owned effective error literal, including reported evaluation-only states.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#CIRCULAR_REF!", "#FUNCTION_NOT_IMPLEMENTED!" ],
    "enumValueDocs" : [ {
      "value" : "#NULL!",
      "summary" : "Excel null-intersection error."
    }, {
      "value" : "#DIV/0!",
      "summary" : "Excel division-by-zero error."
    }, {
      "value" : "#VALUE!",
      "summary" : "Excel wrong-type or wrong-shape value error."
    }, {
      "value" : "#REF!",
      "summary" : "Excel invalid-reference error."
    }, {
      "value" : "#NAME?",
      "summary" : "Excel unknown-name error."
    }, {
      "value" : "#NUM!",
      "summary" : "Excel numeric-domain error."
    }, {
      "value" : "#N/A",
      "summary" : "Excel not-available error."
    }, {
      "value" : "#CIRCULAR_REF!",
      "summary" : "GridGrind-owned circular-reference evaluation error."
    }, {
      "value" : "#FUNCTION_NOT_IMPLEMENTED!",
      "summary" : "GridGrind-owned unsupported-function evaluation error."
    } ]
  } ],
  "targetSelectors" : [ ]
}
```
