<!-- Generated from core metadata; do not edit. -->
# calculationReportType

## CalculationReport

Structured calculation policy, preflight classification, and execution outcome returned on every success and failure response.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| policy | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"calculationPolicySpecType"} |  |
| preflight | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"calculationPreflightReportType"} |  |
| execution | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"calculationExecutionReportType"} |  |

```json
{
  "id" : "CalculationReport",
  "summary" : "Structured calculation policy, preflight classification, and execution outcome returned on every success and failure response.",
  "fields" : [ {
    "name" : "policy",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "calculationPolicySpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "preflight",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "calculationPreflightReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "execution",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "calculationExecutionReportType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
