<!-- Generated from core metadata; do not edit. -->
# executionJournalCalculationType

## ExecutionJournalCalculation

Top-level calculation preflight and execution timings for one request.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| preflight | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"phaseTypes"} |  |
| execution | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"phaseTypes"} |  |

```json
{
  "id" : "ExecutionJournalCalculation",
  "summary" : "Top-level calculation preflight and execution timings for one request.",
  "fields" : [ {
    "name" : "preflight",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "phaseTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "execution",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "phaseTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
