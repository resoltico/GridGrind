<!-- Generated from core metadata; do not edit. -->
# commentAnchorReportType

## CommentAnchorReport

Exact comment-anchor bounds reported from worksheet drawing metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| firstColumn | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| firstRow | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| lastColumn | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| lastRow | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "CommentAnchorReport",
  "summary" : "Exact comment-anchor bounds reported from worksheet drawing metadata.",
  "fields" : [ {
    "name" : "firstColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "firstRow",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "lastColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "lastRow",
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
