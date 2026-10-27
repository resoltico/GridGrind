<!-- Generated from core metadata; do not edit. -->
# commentSpecType

## CommentSpec

Comment payload attached to one cell. Comments can carry ordered rich-text runs and an explicit anchor box.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| text | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |
| author | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| visible | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| runs | OPTIONAL | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"richTextRunSpecType"}} |  |
| anchor | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"commentAnchorSpecType"} |  |

```json
{
  "id" : "CommentSpec",
  "summary" : "Comment payload attached to one cell. Comments can carry ordered rich-text runs and an explicit anchor box.",
  "fields" : [ {
    "name" : "text",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "author",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "visible",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "runs",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "richTextRunSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "anchor",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentAnchorSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
