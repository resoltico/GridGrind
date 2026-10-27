<!-- Generated from core metadata; do not edit. -->
# arrayFormulaReportType

## ArrayFormulaReport

One factual array-formula group report returned by GET_ARRAY_FORMULAS.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| topLeftAddress | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| singleCell | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "ArrayFormulaReport",
  "summary" : "One factual array-formula group report returned by GET_ARRAY_FORMULAS.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "topLeftAddress",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
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
    "name" : "singleCell",
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
