<!-- Generated from core metadata; do not edit. -->
# drawingMarkerReportType

## DrawingMarkerReport

Exact cell-relative drawing marker report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| columnIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":2.147483647E9}]} |  |
| rowIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":2.147483647E9}]} |  |
| dx | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingCoordinateReportTypes"} |  |
| dy | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingCoordinateReportTypes"} |  |

```json
{
  "id" : "DrawingMarkerReport",
  "summary" : "Exact cell-relative drawing marker report.",
  "fields" : [ {
    "name" : "columnIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 2.147483647E9
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "rowIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 2.147483647E9
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "dx",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "drawingCoordinateReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "dy",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "drawingCoordinateReportTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
