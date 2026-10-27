<!-- Generated from core metadata; do not edit. -->
# windowRowReportType

## WindowRowReport

One dense window row preserving row index and ordered cell payloads.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| rowIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"cellReportTypes"}} |  |

```json
{
  "id" : "WindowRowReport",
  "summary" : "One dense window row preserving row index and ordered cell payloads.",
  "fields" : [ {
    "name" : "rowIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "cellReportTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
