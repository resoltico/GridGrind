<!-- Generated from core metadata; do not edit. -->
# calculationExecutionReportType

## CalculationExecutionReport

Post-execution outcome for the authored calculation policy.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| status | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| evaluatedFormulaCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| cachesCleared | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| markRecalculateOnOpenApplied | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| message | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "CalculationExecutionReport",
  "summary" : "Post-execution outcome for the authored calculation policy.",
  "fields" : [ {
    "name" : "status",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NOT_REQUESTED", "SKIPPED", "PARTIAL", "SUCCEEDED", "FAILED" ]
  }, {
    "name" : "evaluatedFormulaCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "cachesCleared",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "markRecalculateOnOpenApplied",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
