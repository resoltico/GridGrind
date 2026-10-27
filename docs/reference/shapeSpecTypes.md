<!-- Generated from core metadata; do not edit. -->
# shapeSpecTypes

## SIMPLE_SHAPE

Write one authored simple shape with required geometry token and optional text payload.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| anchor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingAnchorSpecTypes"} |  |
| presetGeometryToken | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| text | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "SIMPLE_SHAPE",
  "summary" : "Write one authored simple shape with required geometry token and optional text payload.",
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
      "group" : "drawingAnchorSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "presetGeometryToken",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "text",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## CONNECTOR

Write one authored connector with no geometry token or text payload.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| anchor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingAnchorSpecTypes"} |  |

```json
{
  "id" : "CONNECTOR",
  "summary" : "Write one authored connector with no geometry token or text payload.",
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
      "group" : "drawingAnchorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
