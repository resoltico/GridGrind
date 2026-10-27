<!-- Generated from core metadata; do not edit. -->
# executionJournalSpecType

## ExecutionJournalSpec

Explicit execution-journal settings. Use ExecutionJournalSpec.defaults() for SUMMARY detail. SUMMARY returns compact, stable target summaries without timing telemetry, NORMAL returns expanded target summaries with timing telemetry, and VERBOSE also enables fine-grained event emission for host progress rendering.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| level | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} | SUMMARY |

```json
{
  "id" : "ExecutionJournalSpec",
  "summary" : "Explicit execution-journal settings. Use ExecutionJournalSpec.defaults() for SUMMARY detail. SUMMARY returns compact, stable target summaries without timing telemetry, NORMAL returns expanded target summaries with timing telemetry, and VERBOSE also enables fine-grained event emission for host progress rendering.",
  "fields" : [ {
    "name" : "level",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "SUMMARY", "NORMAL", "VERBOSE" ],
    "defaultValue" : "SUMMARY"
  } ],
  "targetSelectors" : [ ]
}
```
