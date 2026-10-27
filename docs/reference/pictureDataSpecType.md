<!-- Generated from core metadata; do not edit. -->
# pictureDataSpecType

## PictureDataSpec

One picture payload with explicit format and base64-encoded binary data.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| format | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| source | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"binarySourceSpecTypes"} |  |

```json
{
  "id" : "PictureDataSpec",
  "summary" : "One picture payload with explicit format and base64-encoded binary data.",
  "fields" : [ {
    "name" : "format",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "EMF", "WMF", "PICT", "JPEG", "PNG", "DIB", "GIF", "TIFF", "EPS", "BMP", "WPG" ]
  }, {
    "name" : "source",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "binarySourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
