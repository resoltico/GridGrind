<!-- Generated from core metadata; do not edit. -->
# commentAnchorSpecType

## CommentAnchorSpec

Explicit comment-anchor bounds measured in zero-based column and row indexes.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| firstColumn | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| firstRow | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| lastColumn | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| lastRow | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "CommentAnchorSpec",
  "summary" : "Explicit comment-anchor bounds measured in zero-based column and row indexes.",
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
