<!-- Generated from core metadata; do not edit. -->
# tableEntryBehaviorReportType

## TableEntryBehaviorReport

Persisted workbook-table behavior toggles.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| hasAutofilter | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| published | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| insertRow | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| insertRowShift | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "TableEntryBehaviorReport",
  "summary" : "Persisted workbook-table behavior toggles.",
  "fields" : [ {
    "name" : "hasAutofilter",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "published",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "insertRow",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "insertRowShift",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
