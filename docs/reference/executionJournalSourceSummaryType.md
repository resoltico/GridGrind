<!-- Generated from core metadata; do not edit. -->
# executionJournalSourceSummaryType

## ExecutionJournalSourceSummary

Journal summary of the authored workbook source.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| type | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| path | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ExecutionJournalSourceSummary",
  "summary" : "Journal summary of the authored workbook source.",
  "fields" : [ {
    "name" : "type",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "path",
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
