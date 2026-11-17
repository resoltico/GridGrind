<!-- Generated from core metadata; do not edit. -->
# drawingAnchorReportTypes

## TWO_CELL

Drawing anchor spanning one start and end marker.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| from | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"drawingMarkerReportType"} |  |
| to | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"drawingMarkerReportType"} |  |
| behavior | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TWO_CELL",
  "summary" : "Drawing anchor spanning one start and end marker.",
  "fields" : [ {
    "name" : "from",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "drawingMarkerReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "to",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "drawingMarkerReportType"
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

## ONE_CELL

Drawing anchor with one start marker plus explicit size.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| from | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"drawingMarkerReportType"} |  |
| widthEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":2.72730423169E13}]} |  |
| heightEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":2.72730423169E13}]} |  |
| behavior | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ONE_CELL",
  "summary" : "Drawing anchor with one start marker plus explicit size.",
  "fields" : [ {
    "name" : "from",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "drawingMarkerReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "widthEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 2.72730423169E13
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "heightEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 2.72730423169E13
      } ]
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

## ABSOLUTE

Drawing anchor with stored absolute coordinates and size.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| x | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingCoordinateReportTypes"} |  |
| y | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingCoordinateReportTypes"} |  |
| widthEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":2.72730423169E13}]} |  |
| heightEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":2.72730423169E13}]} |  |
| behavior | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ABSOLUTE",
  "summary" : "Drawing anchor with stored absolute coordinates and size.",
  "fields" : [ {
    "name" : "x",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "drawingCoordinateReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "y",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "drawingCoordinateReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "widthEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 2.72730423169E13
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "heightEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 2.72730423169E13
      } ]
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
