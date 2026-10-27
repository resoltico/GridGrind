<!-- Generated from core metadata; do not edit. -->
# chartReportType

## ChartReport

One factual chart report with chart-level presentation state and one or more plots.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| anchor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingAnchorReportTypes"} |  |
| title | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"titleTypes"} |  |
| legend | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"legendTypes"} |  |
| displayBlanksAs | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| plotOnlyVisibleCells | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| plots | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"plotTypes"}} |  |

```json
{
  "id" : "ChartReport",
  "summary" : "One factual chart report with chart-level presentation state and one or more plots.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "anchor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "drawingAnchorReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "title",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "titleTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "legend",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "legendTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayBlanksAs",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "GAP", "SPAN", "ZERO" ]
  }, {
    "name" : "plotOnlyVisibleCells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "plots",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "plotTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
