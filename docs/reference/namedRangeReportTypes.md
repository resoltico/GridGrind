<!-- Generated from core metadata; do not edit. -->
# namedRangeReportTypes

## RANGE

Exact named-range report that resolves to one typed workbook target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| scope | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"namedRangeScopeReportTypes"} |  |
| refersToFormula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| target | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"namedRangeTargetTypes"} |  |

```json
{
  "id" : "RANGE",
  "summary" : "Exact named-range report that resolves to one typed workbook target.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "scope",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "namedRangeScopeReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "refersToFormula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "target",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "namedRangeTargetTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## FORMULA

Exact named-range report that remains formula-backed.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| scope | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"namedRangeScopeReportTypes"} |  |
| refersToFormula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FORMULA",
  "summary" : "Exact named-range report that remains formula-backed.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "scope",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "namedRangeScopeReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "refersToFormula",
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
