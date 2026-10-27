<!-- Generated from core metadata; do not edit. -->
# pictureSpecType

## PictureSpec

Named picture-authoring payload for SET_PICTURE.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| image | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"pictureDataSpecType"} |  |
| anchor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingAnchorSpecTypes"} |  |
| description | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "PictureSpec",
  "summary" : "Named picture-authoring payload for SET_PICTURE.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "image",
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
  }, {
    "name" : "description",
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
