<!-- Generated from core metadata; do not edit. -->
# executionJournalFailureStepType

## ExecutionJournalFailureStep

Canonical failing-step reference recorded when an execution failure is attributable to one authored step.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| failedStepIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| failedStepId | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ExecutionJournalFailureStep",
  "summary" : "Canonical failing-step reference recorded when an execution failure is attributable to one authored step.",
  "fields" : [ {
    "name" : "failedStepIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "failedStepId",
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
