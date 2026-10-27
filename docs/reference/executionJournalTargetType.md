<!-- Generated from core metadata; do not edit. -->
# executionJournalTargetType

## ExecutionJournalTarget

One canonical target label recorded inside a step journal.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| kind | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| label | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ExecutionJournalTarget",
  "summary" : "One canonical target label recorded inside a step journal.",
  "fields" : [ {
    "name" : "kind",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "label",
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
