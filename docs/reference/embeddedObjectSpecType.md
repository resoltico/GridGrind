<!-- Generated from core metadata; do not edit. -->
# embeddedObjectSpecType

## EmbeddedObjectSpec

Named embedded-object authoring payload for SET_EMBEDDED_OBJECT. base64Data holds the embedded package bytes and previewImage holds the visible preview raster.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| label | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| fileName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| command | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| payload | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"binarySourceSpecTypes"} |  |
| previewImage | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"pictureDataSpecType"} |  |
| anchor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingAnchorSpecTypes"} |  |

```json
{
  "id" : "EmbeddedObjectSpec",
  "summary" : "Named embedded-object authoring payload for SET_EMBEDDED_OBJECT. base64Data holds the embedded package bytes and previewImage holds the visible preview raster.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "label",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fileName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "command",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "payload",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "binarySourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "previewImage",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "pictureDataSpecType"
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
