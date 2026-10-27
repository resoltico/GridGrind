<!-- Generated from core metadata; do not edit. -->
# headerFooterTextSpecType

## HeaderFooterTextSpec

Plain left, center, and right header or footer text segments. Supply all three fields explicitly.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| left | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |
| center | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |
| right | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "HeaderFooterTextSpec",
  "summary" : "Plain left, center, and right header or footer text segments. Supply all three fields explicitly.",
  "fields" : [ {
    "name" : "left",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "center",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "right",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
