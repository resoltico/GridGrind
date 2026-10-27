<!-- Generated from core metadata; do not edit. -->
# executionJournalTimingType

## ExecutionJournalTiming

Measured timestamps and duration for one execution phase that actually ran.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| startedAt | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| finishedAt | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| durationMillis | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "ExecutionJournalTiming",
  "summary" : "Measured timestamps and duration for one execution phase that actually ran.",
  "fields" : [ {
    "name" : "startedAt",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "finishedAt",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "durationMillis",
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
