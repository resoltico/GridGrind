<!-- Generated from core metadata; do not edit. -->
# calculationPreflightReportType

## CalculationPreflightReport

Formula capability classification captured before server-side evaluation begins.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| scope | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| checkedFormulaCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| summary | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"calculationPreflightSummaryType"} |  |
| formulas | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"formulaCapabilityReportType"}} |  |

```json
{
  "id" : "CalculationPreflightReport",
  "summary" : "Formula capability classification captured before server-side evaluation begins.",
  "fields" : [ {
    "name" : "scope",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "WORKBOOK", "TARGETS" ]
  }, {
    "name" : "checkedFormulaCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "summary",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "calculationPreflightSummaryType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "formulas",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "formulaCapabilityReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
