<!-- Generated from core metadata; do not edit. -->
# tableEntryStructureReportType

## TableEntryStructureReport

Structural table facts including header/totals counts and per-column metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| headerRowCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| totalsRowCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| columnNames | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |
| columns | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"tableColumnReportType"}} |  |

```json
{
  "id" : "TableEntryStructureReport",
  "summary" : "Structural table facts including header/totals counts and per-column metadata.",
  "fields" : [ {
    "name" : "headerRowCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "totalsRowCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "columnNames",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "columns",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "tableColumnReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
