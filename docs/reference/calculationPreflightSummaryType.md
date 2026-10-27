<!-- Generated from core metadata; do not edit. -->
# calculationPreflightSummaryType

## CalculationPreflightSummary

Aggregate counts for evaluable, unevaluable, and unparseable formulas.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| evaluableNowCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| unevaluableNowCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| unparseableByPoiCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "CalculationPreflightSummary",
  "summary" : "Aggregate counts for evaluable, unevaluable, and unparseable formulas.",
  "fields" : [ {
    "name" : "evaluableNowCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "unevaluableNowCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "unparseableByPoiCount",
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
