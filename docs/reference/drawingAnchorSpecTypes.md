<!-- Generated from core metadata; do not edit. -->
# drawingAnchorSpecTypes

## TWO_CELL

Authored two-cell drawing anchor with explicit start and end markers. behavior is explicit on the wire; the Java authoring surface exposes DrawingAnchorSpec.TwoCell.moveAndResize(...).

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| from | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"drawingMarkerSpecType"} |  |
| to | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"drawingMarkerSpecType"} |  |
| behavior | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TWO_CELL",
  "summary" : "Authored two-cell drawing anchor with explicit start and end markers. behavior is explicit on the wire; the Java authoring surface exposes DrawingAnchorSpec.TwoCell.moveAndResize(...).",
  "fields" : [ {
    "name" : "from",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "drawingMarkerSpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "to",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "drawingMarkerSpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "behavior",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "MOVE_AND_RESIZE", "MOVE_DONT_RESIZE", "DONT_MOVE_AND_RESIZE" ]
  } ],
  "targetSelectors" : [ ]
}
```
