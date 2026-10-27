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
| widthEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| heightEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
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
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "heightEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
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

Drawing anchor with absolute EMU coordinates and size.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| xEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| yEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| widthEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| heightEmu | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| behavior | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ABSOLUTE",
  "summary" : "Drawing anchor with absolute EMU coordinates and size.",
  "fields" : [ {
    "name" : "xEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "yEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "widthEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "heightEmu",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
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
